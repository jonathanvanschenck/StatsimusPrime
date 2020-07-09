
from apiclient import errors
from apiclient.http import MediaFileUpload, MediaIoBaseDownload


class IDError(Exception):
    pass

class Service:
    def __init__(self, google_service_object, id = None):
        self.service = google_service_object
        self.id = id

    @property
    def id(self):
        if self.__id is None:
            raise IDError("Service id is uninitialized, use .initialize_env(...)")
        return self.__id
    @id.setter
    def id(self,id):
        self.__id = id

    def __repr__(self):
        return "<Base Service Object>"

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
                fields='nextPageToken, files(id, name, mimeType)',
                pageToken=page_token
            ).execute()
            for file in response.get('files', []):
                children.append({
                    'name':file.get('name'),
                    'id':file.get('id'),
                    'mimeType':file.get('mimeType')
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
        self.service.files().delete(fileId = id)

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


    def copy_to(self, file_id, name, destination_folder_id):
        return self.service.files().copy(
            fileId = file_id,
            body = {
                'name': name,
                'parents': [destination_folder_id]
            }
        ).execute()

class SheetsService(Service):
    def __repr__(self):
        return "<SheetsService Object>"

    def get_values(self, file_id, range):
        return self.service.spreadsheets().values().get(
            spreadsheetId = file_id,
            range = range
        ).execute()

    def update_values(self, file_id, range, values, value_input_option = "USER_ENTERED",
                      major_dimension = "ROWS"):
        # values must be a 2D list, where the outer index matches major_dimension
        return self.service.spreadsheets().values().update(
            spreadsheetId = file_id,
            range = range,
            valueInputOption = value_input_option,
            body = {
                "range" : range,
                "majorDimension" : major_dimension,
                "values" : values
            }
        ).execute()

class StatsService(SheetsService):
    def __repr__(self):
        return "<StatsService Object>"

    def generate_all_values(self):
        for sheet in ['DRAW','IND']:
            yield self.get_values(self.id,"'{}'!A1:ZZ300".format(sheet))

class ScoresheetService(SheetsService):
    def __init__(self, google_service_object, id = None, template_id = None):
        Service.__init__(self, google_service_object, id)
        self.template_id = template_id

    def __repr__(self):
        return "<DriveService Object>"

    @property
    def template_id(self):
        if self.__template_id is None:
            raise IDError("Service template_id is uninitialized, use .initialize_env(...)")
        return self.__template_id
    @template_id.setter
    def template_id(self,id):
        self.__template_id = id

    def generate_all_values(self,file_id):
        for range in ['scoresheet!A1:AD55','metadata!A1:P27','parse_individuals!A1:X140','utils!A1:D26']:
            yield self.get_values(file_id,range)
