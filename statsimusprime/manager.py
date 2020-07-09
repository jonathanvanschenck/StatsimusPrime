# Adapted from https://developers.google.com/docs/api/quickstart/python

import pickle
import os
import json


from urllib.parse import urlparse
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
        self.ss_service = ScoresheetService(ss)

        try:
            self.load_env()
        except FileNotFoundError:
            print("Manager could not find .env file in current directory,\ntry calling .initialize_env(...) to generate it, or\npick a new working directory")
        else:
            print("Manager initialized")

    def __repr__(self):
        return "<Service Object>"

    def initialize_env(self,top_folder_id_or_url):
        top_folder_id = urlparse(top_folder_id_or_url).path.split("/")[-1]

        env = {
            'top_folder_id': top_folder_id,
            'trash_id': None,
            'scoresheets_id': None,
            'stats_id': None,
            'ss_template_id': None
        }


        # Check if any of the files or folders already exist
        files = self.drive_service.get_all_children(top_folder_id)
        for _f in files:
            if _f['mimeType'] == 'application/vnd.google-apps.folder':
                if _f['name'] == 'trash':
                    env['trash_id'] = _f['id']
                    print("Found a trash folder in Google Drive")
                elif _f['name'] == 'scoresheets':
                    env['scoresheets_id'] = _f['id']
                    print("Found a scoresheets folder in Google Drive")
            elif _f['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                if _f['name'] == 'Statistics':
                    env['stats_id'] = _f['id']
                    print("Found a stats document in Google Drive")
                elif _f['name'] == 'Scoresheet_template':
                    env['ss_template_id'] = _f['id']
                    print("Found a scoresheet template in Google Drive")

        # If any don't, create them
        if env['trash_id'] is None:
            env['trash_id'] = self.drive_service.create_folder(
                name='trash',
                parent_folder_id = top_folder_id
            ).get('id')
        if env['scoresheets_id'] is None:
            scoresheets_id = self.drive_service.create_folder(
                name='scoresheets',
                parent_folder_id = top_folder_id
            ).get('id')

        if env['stats_id'] is None:
            stats_id = self.drive_service.upload_excel_as_sheet(
                name = 'Statistics',
                file_path = self.statsfp,
                parent_folder_id = top_folder_id
            ).get('id')

        if env['ss_template_id'] is None:
            ss_template_id = self.drive_service.upload_excel_as_sheet(
                name = 'Scoresheet_template',
                file_path = self.ssfp,
                parent_folder_id = top_folder_id
            ).get('id')

        # create the .env backup
        with open(self.envfp, "w") as f:
            json.dump(env, f, indent=4)

        return self.load_env()

    def load_env(self):
        # Load .env file
        with open(self.envfp) as f:
            env = json.load(f)

        self.top_folder_id = env['top_folder_id']#f.readline().strip().split("=")[1]
        self.trash_id = env['trash_id']#f.readline().strip().split("=")[1]
        self.scoresheets_id = env['scoresheets_id']#f.readline().strip().split("=")[1]
        self.stats_id = env['stats_id']#f.readline().strip().split("=")[1]
        self.ss_template_id = env['ss_template_id']#f.readline().strip().split("=")[1]

        # Activate drive and sheets services
        self.drive_service.id = self.top_folder_id
        self.drive_service.trash_id = self.trash_id

        self.stats_service.id = self.stats_id

        self.ss_service.id = self.scoresheets_id
        self.ss_service.template_id = self.ss_template_id

        return self

    def download_static_image(self,fp=None):
        _fp = fp or os.getcwd()

        # Create a temporary backup folder
        bu_id = self.drive_service.create_folder(
            name = "backup",
            parent_folder_id = self.top_folder_id
        ).get('id')

        print("Copying Statistics")
        # Create a stats backup
        bu_stats_id = self.drive_service.copy_to(
            file_id = self.stats_id,
            name = '_Statistics',
            destination_folder_id = bu_id
        ).get("id")

        # Override all formulas to static values
        for data in self.stats_service.generate_all_values():
            self.stats_service.update_values(
                file_id = bu_stats_id,
                range = data['range'],
                values = data['values']
            )

        print("Copying Scoresheets")
        for ss in self.drive_service.get_all_children(self.scoresheets_id):
            if ss['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                print("Copying",ss['name'])

                # Create a backup
                bu_ss_id = self.drive_service.copy_to(
                    file_id = ss['id'],
                    name = ss['name'],
                    destination_folder_id = bu_id
                ).get("id")

                # Override all formulas to static values
                for data in self.ss_service.generate_all_values(ss['id']):
                    self.ss_service.update_values(
                        file_id = bu_ss_id,
                        range = data['range'],
                        values = data['values']
                    )

        print("Downloading and cleaning")
        for file in self.drive_service.get_all_children(bu_id):
            self.drive_service.download_sheet_as_excel(
                file_id = file['id'],
                destination_file_path = os.path.join(_fp,file['name']+".xlsx")
            )
            self.drive_service.move_to_trash(file['id'])

        self.drive_service.move_to_trash(bu_id)

        return self
