
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
        self.__id = str(id)


    def __repr__(self):
        return "<Base Service Object>"

class DriveService(Service):
    def __init__(self, google_service_object, id = None, trash_id = None):
        Service.__init__(self, google_service_object, id)
        self.trash_id = None

    def __repr__(self):
        return "<DriveService Object>"

    @property
    def trash_id(self):
        if self.__trash_id is None:
            raise IDError("Service id is uninitialized, use .initialize_env(...)")
        return self.__trash_id
    @trash_id.setter
    def trash_id(self,id):
        self.__trash_id = str(id)

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
            body = {'name':'_'+file.get('name')}
            fields='id, parents'
        ).execute()

        return self

    def delete_file(self,id):
        self.service.files().delete(fileId = id)

    def empty_trash(self):
        for file in self.get_all_children(self.trash_id):
            self.delete_file(file['id'])

class StatsService(Service):
    def __repr__(self):
        return "<StatsService Object>"

class ScoresheetService(Service):
    def __repr__(self):
        return "<ScoresheetService Object>"
