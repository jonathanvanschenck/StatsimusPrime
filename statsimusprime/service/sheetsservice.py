import re

from .baseservice import IDError, Service


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
