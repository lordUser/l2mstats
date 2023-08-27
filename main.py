import os

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from PIL import Image
import pytesseract
import pandas as pd

import settings


class L2mStats:
    def __init__(self):
        pytesseract.pytesseract.tesseract_cmd = settings.TESSERACT_PATH
        self.table = self.simple_table

    @property
    def simple_table(self):
        return {'stat': [], 'value': []}

    def make_table(self):
        file_names = os.listdir(settings.IMAGE_DIR)
        text = pytesseract.image_to_string(Image.open(os.path.join(settings.IMAGE_DIR, file_names[0])), lang="rus")
        result = text.split("\n")
        for item in result:
            value = item.split(" ")[-1]
            stat = " ".join(item.split(" ")[:-1])
            if stat:
                self.table['stat'].append(stat)
            if value:
                self.table['value'].append(value)

    def to_csv(self) -> None:
        """
        Method for Creating table and saving it to directory.
        :return: None
        """
        data_frame = self.get_data_frame()
        data_frame.to_csv("table.csv", index=False)

    def get_data_frame(self):
        if self.table == self.simple_table:
            self.make_table()
        return pd.DataFrame(self.table)

    @staticmethod
    def get_creds():
        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', settings.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', settings.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('sheets', 'v4', credentials=creds)

    def create_google_sheet(self):
        service = self.get_creds()
        try:
            # Call the Sheets API
            sheet = service.spreadsheets()
            spreadsheet = sheet.create(body={
                'properties': {'title': 'L2mStats', 'locale': 'ru_RU'},
                'sheets': [{'properties': {'sheetType': 'GRID',
                                           'sheetId': 0,
                                           'title': '',
                                           'gridProperties': {'rowCount': len(self.table['stat']), 'columnCount': 3}}}]
            }).execute()
        except HttpError as err:
            print(err)
        else:
            return spreadsheet["spreadsheetId"]

    def to_google_sheets(self):
        service = self.get_creds()
        sheet_id = self.create_google_sheet()
        df = self.get_data_frame()
        results = service.spreadsheets().values().batchUpdate(spreadsheetId=sheet_id, body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "Лист1!A1:D100",
                 "majorDimension": "ROWS",
                 "values": df.values.tolist()
                 }]
        }).execute()
        print(results["spreadsheetId"])


if __name__ == "__main__":
    l2m_stats = L2mStats()
    l2m_stats.to_google_sheets()
