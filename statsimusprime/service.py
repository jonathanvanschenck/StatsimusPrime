import re

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




LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
class SheetsService(Service):
    def __repr__(self):
        return "<SheetsService Object>"

    @property
    def id(self):
        if self.__id is None:
            raise IDError("Service id is uninitialized, use .initialize_env(...)")
        return self.__id
    @id.setter
    def id(self,id):
        self.__id = id
        if not self.__id is None:
            self.retrieve_spreadsheet_ids()

    def retrieve_spreadsheet_ids(self):
        s = self.service.spreadsheets().get(spreadsheetId = self.__id).execute()

        self.sheet_ids = {
            sheet['properties']['title']: sheet['properties']['sheetId']\
            for sheet in s['sheets']
        }

        self.sheet_properties = {
            sheet['properties']['title']: sheet['properties']\
            for sheet in s['sheets']
        }

        self.named_range_ids = {
            named_range['name'] : named_range['namedRangeId']\
            for named_range in s['namedRanges']
        }

        return self

    def batch_update(self, file_id, request_list):
        # request_list = [
        #   {'update_type': { ... dictionary of update parameters ... }},
        #   ...
        #]
        return self.service.spreadsheets().batchUpdate(
            spreadsheetId = file_id,
            body = {
                "requests" : request_list
            }
        ).execute()


    @staticmethod
    def generate_value_range_json(range, values, major_dim = "ROWS"):
        """Generate the json for a valueRange object

        range : string
            A string in A1 notation which references the cells to be updated

        data : list of lists
            A 2d array (rectangular) which holds the values to be inserted.
        """
        return {
            "range" : range,
            "values" : values,
            "majorDimension" : major_dim
        }


    def batch_update_value(self, file_id, value_range_list, value_input_option = "RAW"):
        """Post a values update request in batch form

        Parameters
        ----------
        file_id : str
            Id of the file to be updated

        value_range_list : list
            A list of valueRange dictionaries to be updated inside the file.
            Note a valueRange dict can be created with `.generate_value_range_json(...)`

        value_input_option : str
            way the values will be input, either "RAW" or "USER_ENTERED"
        """

        return self.service.spreadsheets().values().batchUpdate(
            spreadsheetId = file_id,
            body = {
                "valueInputOption" : value_input_option,
                "data" : value_range_list
            }
        ).execute()

    def batch_get_value(self, file_id, range_list, value_render_option = "FORMATTED_VALUE"):
        """Get values in batch form

        Parameters
        ----------
        file_id : str
            Id of the file to be updated

        range_list : list
            A list of A1 notation ranges to be gotten from the file.

        value_render_option : str
            way the values will be extracted, either "FORMATTED_VALUE" or "FORMULA"
        """

        return self.service.spreadsheets().values().batchGet(
            spreadsheetId = file_id,
            valueRenderOption = value_render_option,
            ranges = range_list
        ).execute()

    def batch_clear_value(self, file_id, range_list):
        """Clear values in batch form

        Parameters
        ----------
        file_id : str
            Id of the file to be updated

        range_list : list
            A list of A1 notation ranges to be cleared from the file.

        """

        return self.service.spreadsheets().values().batchClear(
            spreadsheetId = file_id,
            body = {"ranges": range_list}
        ).execute()

    def batch_copy_over(self, file_id_source, range_list_source,
                        file_id_dest, range_list_dest,
                        value_render_option = "FORMATTED_VALUE",
                        value_input_option = "USER_ENTERED"):
        """Copy ranges between spreadsheet in batch form

        Parameters
        ----------
        file_id_source : str
            The id of the source spreadsheet for the copy.

        range_list_source : list
            List of A1 notation ranges to copy from the source document. Note,
            this MUST be the same length (and each have the same bbox as)
            range_list_dest

        file_id_dest : str
            The id of the destination spreadsheet for the copy

        range_list_dest : list
            List of A1 notation ranges to copy into the destination document. Note,
            this MUST be the same length (and each have the same bbox as)
            range_list_source

        value_render_option : str
            way the values will be extracted, either "FORMATTED_VALUE" or "FORMULA"

        value_input_option : str
            way the values will be input, either "RAW" or "USER_ENTERED"
        """
        response = self.batch_get_value(
            file_id = file_id_source,
            range_list = range_list_source,
            value_render_option = value_render_option
        )

        value_range_list = response['valueRanges']

        for i, range in enumerate(range_list_dest):
            value_range_list[i]['range'] = range

        return self.batch_update_value(
            file_id = file_id_dest,
            value_range_list = value_range_list,
            value_input_option = value_input_option
        )


    @staticmethod
    def get_column_number(column_string):
        """Get the index of the column from a column string

        i.e. A -> 0, AA -> 26
        """
        values = [LETTERS.index(char) for char in column_string[::-1]]
        num = values.pop(0)
        try:
            num += (1+values.pop(0))*(26)
        except IndexError:
            pass

        return num

    @staticmethod
    def generate_bbox_from_A1(range):
        """Generates a bounding box tuple from a range string

        i.e. A1:B3 -> (0,0,3,2), B3:B3 -> (2,1,3,2)
        """
        column1,row1,column2,row2 = re.match("([A-Z]+)(\d+)[:]([A-Z]+)(\d+)",range).groups()
        column1 = SheetsService.get_column_number(column1)
        column2 = SheetsService.get_column_number(column2)

        return (int(row1) - 1, column1, int(row2), column2 + 1)

    @staticmethod
    def generate_A1_from_RC(R,C):
        """Generates a bounding box tuple from a range string

        i.e. A1:B3 -> (0,0,3,2), B3:B3 -> (2,1,3,2)
        """
        a,b = C // 26, C % 26
        if a == 0:
            column = LETTERS[b]
        else:
            column = LETTERS[a - 1] + LETTERS[b]

        return column + str(R + 1)

    @staticmethod
    def generate_grid_range_json(sheet_id, bbox):
        """Generate the json for a gridRange

        Parameters
        ----------
        sheetId : id
            id of the sheet to update inside

        bbox : 4-tuple of ints
            The bounding box of the grid: (row_left, column_top, row_right, column_bottom)
        """
        return {
            "sheetId": sheet_id,
            "startRowIndex": bbox[0],
            "endRowIndex": bbox[2],
            "startColumnIndex": bbox[1],
            "endColumnIndex": bbox[3]
        }

    @staticmethod
    def generate_copy_paste_json(sheet_id, bbox_source, bbox_dest, paste_type = "PASTE_NORMAL"):
        """Generate the json for a copyPaste request

        Parameters
        ----------
        sheetId : id
            id of the sheet to update inside

        bbox_source : 4-tuple of ints
            The bounding box of the grid: (row_left, column_top, row_right, column_bottom)
            for the source of the copy

        bbox-dest : 4-tuple of ints
            The bounding box of the grid: (row_left, column_top, row_right, column_bottom)
            for the destination of the copy

        paste_type : string
            Type of paste, "PASTE_NORMAL", "PASTE_FORMULA", etc
        """
        return {"copyPaste": {
            "source": SheetsService.generate_grid_range_json(sheet_id, bbox_source),
            "destination": SheetsService.generate_grid_range_json(sheet_id, bbox_dest),
            "pasteType": paste_type
        }}

    @staticmethod
    def generate_add_named_range_json(sheet_id, name, bbox):
        """Adds a namedRange to the specified sheet

        Parameters
        ----------
        sheet_id : str
            The id of the sheet to which the namedRange will be attached

        name : str
            The name of the namedRange

        bbox : 4-tuple
            The bounding box of the named range.
        """
        return {"addNamedRange": {
            "namedRange": {
              "name": name,
              "range": SheetsService.generate_grid_range_json(sheet_id, bbox),
            }
          }
        }

    @staticmethod
    def generate_update_named_range_json(sheet_id, named_range_id, name, bbox, fields = "*"):
        """Updates a namedRange in the specified sheet

        Parameters
        ----------
        sheet_id : str
            The id of the sheet to which the namedRange will be updated

        named_range_id : str
            The id of the namedRange

        name : str
            The (new) name of the namedRange

        bbox : 4-tuple
            The (new) bounding box of the named range.

        fields : str
            A fieldMask to specify which fields to update ("*" specifies all)
        """
        return {"updateNamedRange": {
            "namedRange": {
              "namedRangeId": named_range_id,
              "name": name,
              "range": SheetsService.generate_grid_range_json(sheet_id, bbox),
            },
            "fields" : fields
          }
        }

    @staticmethod
    def generate_grid_properties_json(row_count = 1000, column_count = 32,
                                      frozen_row_count = None,
                                      frozen_column_count = None,
                                      hide_gridlines = None):
        """Generates json for a gridProperties object

        Parameters
        ----------
        row_count : int
            Number of rows in the grid

        column_count : int
            Number of columns in the grid

        frozen_row_count : int or None
            Number of frozen rows to include. If None, does not freeze rows.

        frozen_column_count : int or None
            Number of frozen columns to include. If None, does not freeze columns.

        hide_gridlines : Boolean or None
            Indicates if gridlines should be hidden. If None, takes default
        """
        gp_json = {
            "rowCount" : 1000,
            "columnCount" : 32
        }
        if not row_count is None:
            gp_json["rowCount"] = row_count
        if not column_count is None:
            gp_json["columnCount"] = column_count
        if not frozen_row_count is None:
            gp_json["frozenRowCount"] = frozen_row_count
        if not frozen_column_count is None:
            gp_json["frozenColumnCount"] = frozen_column_count
        if not hide_gridlines is None:
            gp_json["hideGridlines"] = hide_gridlines

        return gp_json

    @staticmethod
    def update_sheet_properties_json(sheet_properties, title = None,
                                     grid_properties = None):
        """Takes and exsiting sheetProperties object and returns an updated version

        Parameters
        ----------
        sheet_properties : dict
            A dictionary representing the old sheetProperties object.

        title : str or None
            The new title of the sheet. If None, title will not be updated.

        grid_properties : dict
            The new gridProperties object. If None, it will not be updated,
            Note that gridProperties objects can be created with `.generate_grid_properties_json(...)`
        """
        if not title is None:
            sheet_properties["title"] = title
        if not grid_properties is None:
            sheet_properties["gridProperties"] = grid_properties

        return sheet_properties

    @staticmethod
    def generate_update_sheet_properties_json(sheet_properties, fields = "*"):
        """Updates a sheetProperties in the specified sheet

        Parameters
        ----------
        sheet_properties : dict
            A dictionary representing the (new) sheetProperties object. Can be
            generated using an existing sheetProperties object and `.update_sheet_properties_json()`

        fields : str
            A fieldMask to specify which fields to update ("*" specifies all)
        """
        return {"updateSheetProperties": {
            "properties": sheet_properties,
            "fields" : fields
          }
        }


class StatsService(SheetsService):
    def __repr__(self):
        return "<StatsService Object>"

    @property
    def viewer_id(self):
        if self.__viewer_id is None:
            raise IDError("Service id is uninitialized, use .initialize_env(...)")
        return self.__viewer_id
    @viewer_id.setter
    def viewer_id(self,id):
        self.__viewer_id = id
        if not self.__viewer_id is None:
            self.retrieve_viewer_ids()

    def retrieve_viewer_ids(self):
        s = self.service.spreadsheets().get(spreadsheetId = self.__viewer_id).execute()

        self.viewer_sheet_ids = {
            sheet['properties']['title']: sheet['properties']['sheetId']\
            for sheet in s['sheets']
        }

        self.viewer_sheet_properties = {
            sheet['properties']['title']: sheet['properties']\
            for sheet in s['sheets']
        }

        return self

    def retrieve_meet_parameters(self, roster_json, draw_json):
        self.meet_params = {}
        self.meet_params['total_teams'] = len(set([quizzer['team'] for quizzer in roster_json]))
        self.meet_params['total_quizzes'] = len(draw_json)
        try:
            self.meet_params['prelims_per_team_number'] = 3 * sum([quiz['type'] == "P" for quiz in draw_json]) // self.meet_params['total_teams']
        except ZeroDivisionError:
            self.meet_params['prelims_per_team_number'] = 0
        self.meet_params['total_quizzers'] = len(roster_json)
        try:
            self.meet_params['total_quiz_slots'] = max([int(quiz['slot_num']) for quiz in draw_json])
        except ValueError:
            self.meet_params['total_quiz_slots'] = 0
        try:
            self.meet_params['total_rooms'] = max([int(quiz['room_num']) for quiz in draw_json])
        except ValueError:
            self.meet_params['total_rooms'] = 0

        return self

    # def generate_all_values(self):
    #     for sheet in ['DRAW','IND']:
    #         yield self.get_values(self.id,"'{}'!A1:ZZ300".format(sheet))

    def generate_update_sheet_dimension_json(self, sheet_property_json,
                                             column_count = None,
                                             row_count = None):
        """Generate updateSheetPropertiesRequest to change sheet columnCount/rowCount

        Note, this also updates the `.sheet_properties` object to reflect change

        Parameters
        ----------
        sheet_title : str
            Title of the sheet to be modified

        column_count : int
            The number of columns specified for the sheet

        row_count : int
            The number of rows specified for the sheet. None means do not change
        """

        fields = []
        if not column_count is None:
            fields.append("gridProperties.columnCount")
        if not row_count is None:
            fields.append("gridProperties.rowCount")

        sheet_property_json = self.update_sheet_properties_json(
            sheet_property_json,
            grid_properties = self.generate_grid_properties_json(
                column_count = column_count,
                row_count = row_count
            )
        )
        return self.generate_update_sheet_properties_json(
            sheet_property_json,
            fields = ",".join(fields)
        )

    def set_bracket_weights(self,weights_dictionary):
        """Change the ind score weighting of the bracket

        weights : dictionary
            Dictionary of weight values (floats, ints or None), with any of the following
            keys: ["P", "S", "A", "B"]. If "None", then that bracket type does not
            contribute to the total weight.
        """

        processed_weights = [[1.0],[1.0],[0.7],[0.5]]
        for i,k in enumerate("PSAB"):
            try:
                w = weights_dictionary.pop(k)
            except KeyError:
                pass
            else:
                if w is None:
                    processed_weights[i][0] = "NA"
                else:
                    processed_weights[i][0] = w

        value_range_list = [self.generate_value_range_json(
            range = "Utils!C2:C5",
            values = processed_weights
        )]

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list
        )

        return self


    def set_roster(self,roster_json):
        """Update the contents of the roster

        Parameters
        ----------
        roster_json : list
            list of dictionaries representing each quizzer. Each dictionary should
            have the keys: ["id", "team", "bib", "name", "moniker", "is_rookie", "is_cap", "is_cc"]
        """

        value_range_list = []

        column_names = {"id":0, "name":1, "moniker":2, "is_rookie":3, "is_cap":4, "is_cc":5}

        team_list = sorted(list(set([quizzer["team"] for quizzer in roster_json])))
        roster_matrix = [[team] + (5 * 6) * [""] for team in team_list]

        for ti, team in enumerate(team_list):
            for quizzer in [quizzer for quizzer in roster_json if quizzer["team"] == team]:
                offset = 6 * (int(quizzer['bib']) - 1)
                if roster_matrix[ti][1 + offset] != "":
                    # Log if quizzer is overwritten
                    print("Bib numbering error, both {} and {} have bib {}".format(
                        roster_matrix[ti][1 + offset + column_names['name']],
                        quizzer['name'],
                        quizzer['bib']
                    ))
                for k,v in column_names.items():
                    roster_matrix[ti][1 + offset + v] = quizzer[k]


        value_range_list.append(
            self.generate_value_range_json(
                range = "Roster!A3:AE" + str(len(team_list) + 3),
                values = roster_matrix
            )
        )

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list
        )

        return self

    def set_draw(self, draw_json):#, roster_json):
        """Update the contents of the draw

        Parameters
        ----------
        draw_json : list
            list of dictionaries representing each quiz. Each dictionary should
            have the keys: ["quiz_num","slot_num","room_num","slot_time":,"team1","team2","team3", "url", "type"]

        roster_json : list
            list of dictionaries representing each quizzer. Each dictionary should
            have the keys: ["id", "team", "bib", "name", "moniker", "is_rookie", "is_cap", "is_cc"]
        """


        # Step 1: Insert draw into DrawLookup
        column_names_left = ["quiz_num","slot_time","room_num","slot_num"]
        column_names_right = ["team1","team2","team3"]

        draw_matrix_left = [[quiz[key] for key in column_names_left] for quiz in draw_json]
        draw_matrix_right = []
        for quiz in draw_json:
            if "_" in quiz['team1']:
                if quiz['team1'][0] == "P":
                    # Calculate post-prelim ranking lookup
                    quiz_row = []
                    for key in column_names_right:
                        rank = int(quiz[key].split("_")[-1])
                        lookup = "TeamSummary!B{}".format(2+rank)
                        quiz_row.append('={}'.format(lookup))
                else:
                    # Calculate schedule lookup
                    quiz_row = []
                    for key in column_names_right:
                        quiz_num, placement = quiz[key].split("_")
                        quiz_previous = [q for q in draw_json if q['quiz_num'] == quiz_num][0]
                        offset_row = 2 + 3 * (int(quiz_previous['slot_num']) - 1)
                        offset_column = 2 + 4 * (int(quiz_previous['room_num']) - 1)
                        team_range = "Schedule!{}:{}".format(
                            self.generate_A1_from_RC(offset_row + 0, offset_column + 1),
                            self.generate_A1_from_RC(offset_row + 2, offset_column + 1)
                        )
                        placement_range = "Schedule!{}:{}".format(
                            self.generate_A1_from_RC(offset_row + 0, offset_column + 3),
                            self.generate_A1_from_RC(offset_row + 2, offset_column + 3)
                        )
                        error_msg = "{placement}{suffix} in {quiz_num}".format(
                            placement = placement,
                            suffix = {"1":"st","2":"nd","3":"rd"}[placement],
                            quiz_num = quiz_num
                        )

                        quiz_row.append(
                            '=IFERROR(INDEX({team_range},MATCH({placement},{placement_range},0),0),"{error_msg}")'.format(
                                team_range = team_range,
                                placement = placement,
                                placement_range = placement_range,
                                error_msg = error_msg
                            )
                        )
            else:
                # Just add the prelim quiz
                quiz_row = [quiz[key] for key in column_names_right]

            draw_matrix_right.append(quiz_row)
        #draw_matrix_right = [[quiz[key] for key in column_names_right] for quiz in draw_json]

        value_range_list = [
            self.generate_value_range_json(
                range = "DrawLookup!B3:E" + str(len(draw_matrix_left) + 3),
                values = draw_matrix_left
            ),
            self.generate_value_range_json(
                range = "DrawLookup!G3:I" + str(len(draw_matrix_right) + 3),
                values = draw_matrix_right
            )
        ]
        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )


        # Step 2: Prepare the rest of the DrawLookup page
        #TN = self.meet_params['total_teams']#len(set([quizzer['team'] for quizzer in roster_json]))
        #TQN = self.meet_params['total_quizzes']#len(draw_json)
        sheet_id = self.sheet_ids['DrawLookup']

        requests = []

        # Set sheet width to 11 + 'total_teams'
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.sheet_properties['DrawLookup'],
            column_count = 11 + self.meet_params['total_teams']
        ))

        # Copy L2 right 'total_teams' times
        bbox_source = list(self.generate_bbox_from_A1("L2:L2"))
        bbox_dest = 1*bbox_source
        bbox_dest[1] += 1
        bbox_dest[3] += self.meet_params['total_teams'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy L3 right 'total_teams' times and down 'total_quizzes' times
        bbox_source = list(self.generate_bbox_from_A1("L3:L3"))
        bbox_dest = 1*bbox_source
        bbox_dest[2] += self.meet_params['total_quizzes'] - 1
        bbox_dest[3] += self.meet_params['total_teams'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy F3 down 'total_quizzes' times
        bbox_source = list(self.generate_bbox_from_A1("F3:F3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += self.meet_params['total_quizzes'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy K3 down 'total_quizzes' times
        bbox_source = list(self.generate_bbox_from_A1("K3:K3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += self.meet_params['total_quizzes'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Update QUIZINDEXLOOKUP to be DrawLookup!L3:(L3+TN+TQN)
        bbox = list(self.generate_bbox_from_A1("L3:L3"))
        bbox[2] += self.meet_params['total_quizzes'] - 1
        bbox[3] += self.meet_params['total_teams'] - 1
        requests.append(
            self.generate_update_named_range_json(
                sheet_id = sheet_id,
                named_range_id = self.named_range_ids['QUIZINDEXLOOKUP'],
                name = "QUIZINDEXLOOKUP",
                bbox = bbox,
                fields = "range"
            )
        )

        response = self.batch_update(
            file_id = self.id,
            request_list = requests
        )

        return self

    def initialize_team_summary(self):#, roster_json):
        """Prepares the TeamSummary tab

        The copies down columns E and F
        """
        #TN = len(set([quizzer['team'] for quizzer in roster_json]))

        sheet_id = self.sheet_ids['TeamSummary']
        requests = []

        # Set sheet width to 10
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.sheet_properties['TeamSummary'],
            column_count = 10
        ))

        # Copy down E3:F3
        bbox_source = list(self.generate_bbox_from_A1("E3:F3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += self.meet_params['total_teams'] - 1
        requests.append(self.generate_copy_paste_json(
            sheet_id = sheet_id,
            bbox_source = bbox_source,
            bbox_dest = bbox_dest
        ))

        response = self.batch_update(
            file_id = self.id,
            request_list = requests
        )

        return self

    def initialize_schedule(self):#, draw_json):
        """Prepares the schedule tab
        """

        #TSN = max([int(quiz['slot_num']) for quiz in draw_json])
        #TRN = max([int(quiz['room_num']) for quiz in draw_json])

        sheet_id = self.sheet_ids['Schedule']

        requests = []

        # Set sheet width to 2 + 4*TRN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.sheet_properties['Schedule'],
            column_count = 2 + 4*self.meet_params['total_rooms'],
            row_count = 2 + 3*self.meet_params['total_quiz_slots']
        ))

        # Copy down A3:F5
        bbox_source = list(self.generate_bbox_from_A1("A3:F5"))
        bbox_dest = 1*bbox_source
        for i in range(1,self.meet_params['total_quiz_slots']):
            # Shift window down 3 rows
            bbox_dest[0] += 3
            bbox_dest[2] += 3
            requests.append(self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest
            ))

        # Copy right C1:F
        bbox_source = list(self.generate_bbox_from_A1("C1:F"+str(2+3*self.meet_params['total_quiz_slots'])))
        bbox_dest = 1*bbox_source
        for i in range(1,self.meet_params['total_rooms']):
            # Shift the window right 4 columns
            bbox_dest[1] += 4
            bbox_dest[3] += 4
            requests.append(self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest
            ))

        self.batch_update(
            file_id = self.id,
            request_list = requests
        )

        return self


    def set_team_parsed(self):#, draw_json, roster_json):
        #TN = len(set([quizzer['team'] for quizzer in roster_json]))
        #PN = 3 * sum([quiz['type'] == "P" for quiz in draw_json]) // TN

        # Update Quiz Total and Average team points formulas
        points_cell_string = ", ".join([
            self.generate_A1_from_RC(2, 8 + i * 5) for i in range(self.meet_params['prelims_per_team_number'])
        ])
        value_range_list = [
            self.generate_value_range_json(
                range = "TeamParsed!C3:C3",
                values = [["=SUM({})".format(points_cell_string)]]
            ),
            self.generate_value_range_json(
                range = "TeamParsed!D3:D3",
                values = [['=IFERROR(AVERAGE({}), 0)'.format(points_cell_string)]]
            )
        ]

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )



        sheet_id = self.sheet_ids['TeamParsed']
        requests = []

        # Set sheet width to 4 + 5*PN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.sheet_properties['TeamParsed'],
            column_count = 4 + 5*self.meet_params['prelims_per_team_number']
        ))

        # Copy A3 down TN times
        bbox_source = list(self.generate_bbox_from_A1("A3:A3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += self.meet_params['total_teams'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy C3:I3 down TN times
        bbox_source = list(self.generate_bbox_from_A1("C3:I3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += self.meet_params['total_teams'] - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy E1:(I2+TN) right PN times
        bbox_source = list(self.generate_bbox_from_A1("E1:I1"))
        bbox_source[2] += 1 + self.meet_params['total_teams'] # add header and TN teams
        bbox_dest = 1*bbox_source
        for i in range(1,self.meet_params['prelims_per_team_number']):
            # Shift window right 5
            bbox_dest[1] += 5
            bbox_dest[3] += 5
            requests.append(
                self.generate_copy_paste_json(
                    sheet_id = sheet_id,
                    bbox_source = bbox_source,
                    bbox_dest = bbox_dest,
                    paste_type = "PASTE_NORMAL"
                )
            )


        self.batch_update(
            file_id = self.id,
            request_list = requests
        )

        return self

    def set_individual_parsed(self, roster_json):
        #TN = len(set([quizzer['team'] for quizzer in roster_json]))
        #PN = 3 * sum([quiz['type'] == "P" for quiz in draw_json]) // TN
        #QN = len(roster_json)

        value_range_list = []

        # Step 1: inject roster
        column_names = ["id", "name", "moniker", "team", "bib"]
        roster_matrix = [[quizzer[k] for k in column_names] for quizzer in roster_json]
        QN = len(roster_matrix)

        value_range_list.append(
            self.generate_value_range_json(
                range = "IndividualParsed!B3:F" + str(QN + 3),
                values = roster_matrix
            )
        )

        # Step 2: correct formulas
        column_indicies = [8 + 11 + i*12 for i in range(self.meet_params['prelims_per_team_number']+4)] # prelims
        points_cell_string = ", ".join([
            self.generate_A1_from_RC(2, ci) for ci in column_indicies
        ])
        value_range_list.append(
            self.generate_value_range_json(
                range = "IndividualParsed!G3:G3",
                values = [["=IFERROR(AVERAGE({}), 0)".format(points_cell_string)]]
            )
        )
        value_range_list.append(
            self.generate_value_range_json(
                range = "IndividualParsed!T3:T3",
                values = [['=IF(S3="","",S3*INDEX(WEIGHTSVALUE,MATCH(IF(I$1<={},"P",$H3),WEIGHTSKEY,0),0))'.format(self.meet_params['prelims_per_team_number'])]]
            )
        )

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )


        sheet_id = self.sheet_ids['IndividualParsed']
        requests = []

        # Set sheet width to 8 + 12*(PN+4)
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.sheet_properties['IndividualParsed'],
            column_count = 8 + 12*(self.meet_params['prelims_per_team_number'] + 4)
        ))

        # Copy A3 down QN times
        bbox_source = list(self.generate_bbox_from_A1("A3:A3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += QN - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )

        # Copy G3:T3 down QN times
        bbox_source = list(self.generate_bbox_from_A1("G3:T3"))
        bbox_dest = 1*bbox_source
        bbox_dest[0] += 1
        bbox_dest[2] += QN - 1
        requests.append(
            self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest,
                paste_type = "PASTE_FORMULA"
            )
        )


        # Header, P1,  P2, ... , PN ,  B1 , B2 , B3 , B4
        #   8     12   12  ...   12    12   12   12   12

        # Copy quiz section over:
        bbox_source = list(self.generate_bbox_from_A1("I1:T1"))
        bbox_source[2] += 1 + QN # add header and QN quizzers
        bbox_dest = 1*bbox_source
        for i in range(1, self.meet_params['prelims_per_team_number'] + 4):
            # Shift window right 12
            bbox_dest[1] += 12
            bbox_dest[3] += 12
            requests.append(
                self.generate_copy_paste_json(
                    sheet_id = sheet_id,
                    bbox_source = bbox_source,
                    bbox_dest = bbox_dest,
                    paste_type = "PASTE_NORMAL"
                )
            )


        self.batch_update(
            file_id = self.id,
            request_list = requests
        )

        return self

    def initialize_viewer(self):
        """Prepares the viewer's schedule tab
        """

        # # Clean the viewer
        # range_list = []
        #
        # range_list.append("Schedule!B3:E5")
        #
        # self.batch_clear_value(
        #     file_id = self.viewer_id,
        #     range_list = range_list
        # )


        sheet_id = self.viewer_sheet_ids['Schedule']

        requests = []

        # Set Schedule sheet width to 2 + 4*TRN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.viewer_sheet_properties['Schedule'],
            column_count = 2 + 4*self.meet_params['total_rooms'],
            row_count = 2 + 3*self.meet_params['total_quiz_slots']
        ))
        # Set DrawLookup sheet to 8 by 2+TQN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.viewer_sheet_properties['DrawLookup'],
            column_count = 8,
            row_count = 2 + self.meet_params['total_quizzes']
        ))
        # Set Roster sheet to 4*5 + 1 by 2+TN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.viewer_sheet_properties['Roster'],
            column_count = 1 + 4*5,
            row_count = 2 + self.meet_params['total_teams']
        ))
        # Set TeamSummary sheet to rows 2+TN
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.viewer_sheet_properties['TeamSummary'],
            row_count = 2 + self.meet_params['total_teams']
        ))
        # Set IndividualSummary sheet to rows 2+QN+5
        requests.append(self.generate_update_sheet_dimension_json(
            sheet_property_json = self.viewer_sheet_properties['IndividualSummary'],
            row_count = 2 + self.meet_params['total_quizzers'] + 5
        ))


        # Copy down A3:F5
        bbox_source = list(self.generate_bbox_from_A1("A3:F5"))
        bbox_dest = 1*bbox_source
        for i in range(1,self.meet_params['total_quiz_slots']):
            # Shift window down 3 rows
            bbox_dest[0] += 3
            bbox_dest[2] += 3
            requests.append(self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest
            ))

        # Copy right C1:F
        bbox_source = list(self.generate_bbox_from_A1("C1:F"+str(2+3*self.meet_params['total_quiz_slots'])))
        bbox_dest = 1*bbox_source
        for i in range(1,self.meet_params['total_rooms']):
            # Shift the window right 4 columns
            bbox_dest[1] += 4
            bbox_dest[3] += 4
            requests.append(self.generate_copy_paste_json(
                sheet_id = sheet_id,
                bbox_source = bbox_source,
                bbox_dest = bbox_dest
            ))

        self.batch_update(
            file_id = self.viewer_id,
            request_list = requests
        )

        return self

    # def copy_over_all(self):
    #     """Copies everything from stats doc to viewer doc
    #     """
    #     return self.copy_over_draw()

    def copy_over_schedule(self):
        """Copies the schedule tab from stats to viewer
        """

        str_temp = "Schedule!{}:{}"
        range_list_source = []
        range_list_dest = []

        # Copy B3:?? to A3:??
        range_list_source.append(str_temp.format(
            self.generate_A1_from_RC(2,1),
            self.generate_A1_from_RC(
                2 + 3*self.meet_params['total_quiz_slots'] - 1,
                2 + 4*self.meet_params['total_rooms'] - 1
            )
        ))
        range_list_dest.append(str_temp.format(
            self.generate_A1_from_RC(2,0),
            self.generate_A1_from_RC(
                2 + 3*self.meet_params['total_quiz_slots'] - 1,
                1 + 4*self.meet_params['total_rooms'] - 1
            )
        ))


        self.batch_copy_over(
            file_id_source = self.id,
            range_list_source = range_list_source,
            file_id_dest = self.viewer_id,
            range_list_dest = range_list_dest,
            value_render_option = "FORMATTED_VALUE",
            value_input_option = "USER_ENTERED"
        )

        return self

    def copy_over_draw(self):
        """Copies the DrawLookup tab from stats to viewer
        """

        str_temp = "DrawLookup!{}:{}"
        range_list_source = []
        range_list_dest = []

        # Copy B3:I(2+TQN) to A3:H(2+TQN)
        range_list_source.append(str_temp.format(
            "B3",
            "I"+str(2+self.meet_params['total_quizzes'])
        ))
        range_list_dest.append(str_temp.format(
            "A3",
            "H"+str(2+self.meet_params['total_quizzes'])
        ))


        self.batch_copy_over(
            file_id_source = self.id,
            range_list_source = range_list_source,
            file_id_dest = self.viewer_id,
            range_list_dest = range_list_dest,
            value_render_option = "FORMATTED_VALUE",
            value_input_option = "USER_ENTERED"
        )

        return self

    def copy_over_roster(self):
        """Copies the DrawLookup tab from stats to viewer
        """

        str_temp = "Roster!{}:{}"
        range_list_source = []
        range_list_dest = []

        # Copy A3:A(2+TN) to A3:A(2+TN)
        range_list_source.append(str_temp.format(
            "A3",
            "A"+str(2+self.meet_params['total_teams'])
        ))
        range_list_dest.append(str_temp.format(
            "A3",
            "A"+str(2+self.meet_params['total_teams'])
        ))

        # Copy last four columns of each window over
        for i in range(5):
            range_list_source.append(str_temp.format(
                self.generate_A1_from_RC(
                    2,
                    1 + 2 + 6*i
                ),
                self.generate_A1_from_RC(
                    2 + self.meet_params['total_teams'] - 1,
                    1 + 2 + 6*i + 4 - 1
                )
            ))
            range_list_dest.append(str_temp.format(
                self.generate_A1_from_RC(
                    2,
                    1 + 4*i
                ),
                self.generate_A1_from_RC(
                    2 + self.meet_params['total_teams'] - 1,
                    1 + 4*i + 4 - 1
                )
            ))


        self.batch_copy_over(
            file_id_source = self.id,
            range_list_source = range_list_source,
            file_id_dest = self.viewer_id,
            range_list_dest = range_list_dest,
            value_render_option = "FORMATTED_VALUE",
            value_input_option = "USER_ENTERED"
        )

        return self

    def copy_over_team_summary(self):
        """Copies the TeamSummary tab from stats to viewer
        """

        str_temp = "TeamSummary!{}:{}"
        range_list_source = []
        range_list_dest = []

        # Copy A3:I(2+TN) to A3:I(2+TN)
        range_list_source.append(str_temp.format(
            "A3",
            "I"+str(2+self.meet_params['total_teams'])
        ))
        range_list_dest.append(str_temp.format(
            "A3",
            "I"+str(2+self.meet_params['total_teams'])
        ))


        self.batch_copy_over(
            file_id_source = self.id,
            range_list_source = range_list_source,
            file_id_dest = self.viewer_id,
            range_list_dest = range_list_dest,
            value_render_option = "FORMATTED_VALUE",
            value_input_option = "USER_ENTERED"
        )

        return self

    def copy_over_individual_summary(self):
        """Copies the IndividualSummary tab from stats to viewer
        """

        str_temp = "IndividualSummary!{}:{}"
        range_list_source = []
        range_list_dest = []

        # Copy A3:A(2+QN) to A3:A(2+QN)
        range_list_source.append(str_temp.format(
            "A3",
            "A"+str(2+self.meet_params['total_quizzers'])
        ))
        range_list_dest.append(str_temp.format(
            "A3",
            "A"+str(2+self.meet_params['total_quizzers'])
        ))

        # Copy D3:H(2+QN) to B3:F(2+QN)
        range_list_source.append(str_temp.format(
            "D3",
            "H"+str(2+self.meet_params['total_quizzers'])
        ))
        range_list_dest.append(str_temp.format(
            "B3",
            "F"+str(2+self.meet_params['total_quizzers'])
        ))


        self.batch_copy_over(
            file_id_source = self.id,
            range_list_source = range_list_source,
            file_id_dest = self.viewer_id,
            range_list_dest = range_list_dest,
            value_render_option = "FORMATTED_VALUE",
            value_input_option = "USER_ENTERED"
        )

        return self

    # def remove_ss_urls(self):
    #     """Remove url references to scoresheets from DrawLookup
    #     """
    #     self.batch_clear_value(
    #         file_id = self.viewer_id,
    #         range_list = ["DrawLookup!J3:J"+str(3+self.meet_params['total_quizzes'])]
    #     )
    #     return self

    def update_ss_urls(self,draw_json):
        """Updates the url references to scoresheets in DrawLookup
        """

        values = [['="{}"'.format(quiz['url'])] for quiz in draw_json]

        value_range_list = [
            self.generate_value_range_json(
                range = "DrawLookup!J3:J"+str(3+self.meet_params['total_quizzes']-1),
                values = values
            )
        ]

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )

        return self

class ScoresheetService(SheetsService):
    def __repr__(self):
        return "<ScoresheetService Object>"

    def initialize_global_variables(self, viewer_url):
        """Initializes the global variables for the scoresheet template

        viewer_url : str
            The share url for the viewer document
        """

        values = [
            ['="{}"'.format(viewer_url)], # Viewer URL
            ['="Roster!$A$3:$U$100"'], # Roster Range
            ['="DrawLookup!$A$3:$H$300"'], # Draw Range
            [''] # TODO: allow officials imports
        ]

        value_range_list = [
            self.generate_value_range_json(
                range = "utils!B2:B5",
                values = values
            )
        ]

        self.batch_update_value(
            file_id = self.id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )


        return self

    def set_quiz_number_for(self, file_id, quiz_num):
        """Sets the quiz number for a copy of the template

        file_id : str
            The id of the template copy to have its quiz number set

        quiz_num : str
            The number of the quiz
        """

        value_range_list = [
            self.generate_value_range_json(
                range = "metadata!B16:B16",
                values = [[quiz_num]]
            )
        ]

        self.batch_update_value(
            file_id = file_id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )

        return self
