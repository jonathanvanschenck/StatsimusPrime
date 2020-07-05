# Adapted from https://developers.google.com/docs/api/quickstart/python

import pickle
import os

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from apiclient.http import MediaFileUpload

from statsimusprime.service import DriveService, StatsService, ScoresheetService

SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']


class Manager:
    def __init__(self,working_directory=None):
        if working_directory is None:
            wd = os.getcwd()
        else:
            wd = working_directory
        tokenfp = os.path.join(wd,'token.pickle')
        credfp = os.path.join(wd,'credentials.json')
        self.ssfp = os.path.join(wd,'templates','Scoresheet_template.xlsx')
        self.statsfp = os.path.join(wd,'templates','Statistics_template.xlsx')
        self.envfp = os.path.join(wd,'.env')
        creds = None
        if os.path.exists(tokenfp):
            with open(tokenfp, 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                            credfp,
                            SCOPES
                        )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(tokenfp, 'wb') as token:
                pickle.dump(creds, token)

        self.drive_service = DriveService(build('drive', 'v3', credentials=creds))
        ss = build('sheets', 'v4', credentials=creds)
        self.stats_service = StatsService(ss)
        self.ss_serice = ScoresheetService(ss)

        try:
            self.load_env()
        except FileNotFoundError:
            print("Manager could not find .env file in current directory,\
                   try calling .initialize_env(...) to generate it, or\
                   pick a new working directory")
        else:
            print("Manager initialized")

    def __repr__(self):
        return "<Service Object>"

    def initialize_env(self,top_folder_id):
        files = self.get_all_children(top_folder_id)

        trash_id = None
        scoresheets_id = None
        stats_id = None
        ss_template_id = None
        for _f in files:
            if _f['mimeType'] == 'application/vnd.google-apps.folder':
                if _f['name'] == 'trash':
                    trash_id = _f['id']
                elif _f['name'] == 'scoresheets':
                    scoresheets_id = _f['id']
            elif _f['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                if _f['name'] == 'Statistics':
                    stats_id = _f['id']
                elif _f['name'] == 'Scoresheet_template':
                    ss_template_id = _f['id']

        file_metadata = {
            'name': '',
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [top_folder_id]
        }
        if trash_id is None:
            file_metadata['name'] = 'trash'
            _f = self.drive_service.files().create(
                body = file_metadata,
                fields = 'id'
            ).execute()
            trash_id = _f.get('id')
        if scoresheets_id is None:
            file_metadata['name'] = 'scoresheets'
            _f = self.drive_service.files().create(
                body = file_metadata,
                fields = 'id'
            ).execute()
            scoresheets_id = _f.get('id')

        file_metadata = {
            'name': '',
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [top_folder_id]
        }
        if stats_id is None:
            file_metadata['name'] = 'Statistics'
            media = MediaFileUpload(
                self.statsfp,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                resumable=True
            )
            _f = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            stats_id = _f.get('id')
        if ss_template_id is None:
            file_metadata['name'] = 'Scoresheet_template'
            media = MediaFileUpload(
                self.ssfp,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                resumable=True
            )
            _f = self.drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            ss_template_id = _f.get('id')


        with open(self.envfp, "w") as f:
            f.write("top_folder_id={}\n".format(top_folder_id))
            f.write("trash_id={}\n".format(trash_id))
            f.write("scoresheets_id={}\n".format(scoresheets_id))
            f.write("stats_id={}\n".format(stats_id))
            f.write("ss_template_id={}\n".format(ss_template_id))

        self.load_env()

        return self

    def load_env(self):
        with open(self.envfp) as f:
            self.top_folder_id = f.readline().strip().split("=")[1]
            self.trash_id = f.readline().strip().split("=")[1]
            self.scoresheets_id = f.readline().strip().split("=")[1]
            self.stats_id = f.readline().strip().split("=")[1]
            self.ss_template_id = f.readline().strip().split("=")[1]

        return self
