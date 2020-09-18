
from apiclient.http import MediaFileUpload, MediaIoBaseDownload

from .baseservice import IDError, Service


class DriveService(Service):
    def __init__(self, google_service_object, id = None, trash_id = None):
        Service.__init__(self, google_service_object, id)
        self.trash_id = trash_id

    def __repr__(self):
        return "<DriveService Object>"

    @property
    def trash_id(self):
        if self.__trash_id is None:
            raise IDError("Service trash_id is uninitialized, use .initialize_env(...)")
        return self.__trash_id
    @trash_id.setter
    def trash_id(self,id):
        self.__trash_id = id

    def get_all_children(self,folder_id):
        # adapted from https://developers.google.com/drive/api/v3/search-files
        page_token = None
        children = []
        while True:
            response = self.service.files().list(
                q="'{}' in parents".format(folder_id),
                spaces='drive',
                fields='nextPageToken, files(id, name, mimeType, parents)',
                pageToken=page_token
            ).execute()
            for file in response.get('files', []):
                children.append({
                    'name':file.get('name'),
                    'id':file.get('id'),
                    'mimeType':file.get('mimeType'),
                    'parents':file.get('parents')
                })
            page_token = response.get('nextPageToken', None)
            if page_token is None:
                break
        return children

    def move_to_trash(self,file_id):
        # Get file's current parents
        file = self.service.files().get(
            fileId=file_id,
            fields='name, parents'
        ).execute()
        # Remove old parents and attach new parent "trash"
        previous_parents = ",".join(file.get('parents'))
        file = self.service.files().update(
            fileId=file_id,
            addParents=self.trash_id,
            removeParents=previous_parents,
            body = {'name':'_'+file.get('name')},
            fields='id, parents'
        ).execute()

        return self

    def delete_file(self,id):
        self.service.files().delete(fileId = id).execute()

        return self

    def delete_recursive(self, id, verbose = True):
        file = self.service.files().get(fileId = id).execute()
        self.__delete_recusive(file, verbose)

        return self

    def __delete_recusive(self, file, verbose = True):
        if file['mimeType'] == 'application/vnd.google-apps.folder':
            for _file in self.get_all_children(file['id']):
                self.__delete_recusive(_file)
        if verbose:
            print("Deleting: ",file['name'])
        self.delete_file(file['id'])

        return

    def empty_trash(self):
        for file in self.get_all_children(self.trash_id):
            self.delete_file(file['id'])

    def create_folder(self,name,parent_folder_id=None):
        pfid = parent_folder_id or ""
        file_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [pfid]
        }
        return self.service.files().create(
            body = file_metadata,
            fields = 'id'
        ).execute()

    def upload_excel_as_sheet(self,name,file_path,parent_folder_id=None):
        pfid = parent_folder_id or ""
        file_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [pfid]
        }
        media = MediaFileUpload(
            file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            resumable=True
        )
        return self.service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

    def download_sheet_as_excel(self, file_id, destination_file_path, verbose = False):
        request = self.service.files().export_media(
            fileId = file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        with open(destination_file_path,"wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                if verbose:
                    print("Download %d%%." % int(status.progress() * 100))

        return self


    def upload_json(self,name,file_path,parent_folder_id=None):
        """Uploads a new json file into the drive
        """
        pfid = parent_folder_id or ""

        file_metadata = {
            'name': name,
            'mimeType': 'application/json',
            'parents': [pfid]
        }

        media = MediaFileUpload(
            file_path,
            mimetype='application/json',
            resumable=True
        )

        return self.service.files().create(
            body=file_metadata,
            media_body=media,
            fields = "id"
        ).execute()


    def update_json(self, file_id, file_path):
        """Updates an existing json file in the drive
        """
        media = MediaFileUpload(
            file_path,
            mimetype='application/json',
            resumable=True
        )

        self.service.files().update(
            fileId = file_id,
            media_body = media
        ).execute()

        return self

    def download_json(self, file_id, destination_file_path, verbose = False):
        """Downloads a json file
        """
        # request = self.service.files().export_media(
        #     fileId = file_id,
        #     mimeType='application/json'
        # )
        request = self.service.files().get_media(
            fileId = file_id
        )
        with open(destination_file_path,"wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                if verbose:
                    print("Download %d%%." % int(status.progress() * 100))

        return self


    def copy_to(self, file_id, name, destination_folder_id, fields = "id"):
        return self.service.files().copy(
            fileId = file_id,
            fields = fields,
            body = {
                'name': name,
                'parents': [destination_folder_id]
            }
        ).execute()


    def get_file_url(self, file_id):
        return self.service.files().get(
            fileId = file_id,
            fields = "webViewLink"
        ).execute().get("webViewLink")

    def publish_file(self, file_id):
        """Add a 'anyoneWithLinkCanView' permission to the file
        """
        return self.service.permissions().create(
            fileId = file_id,
            body = {
                "role": "reader",
                "type": "anyone"
            }
        ).execute()
