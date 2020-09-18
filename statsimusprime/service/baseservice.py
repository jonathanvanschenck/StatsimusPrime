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
