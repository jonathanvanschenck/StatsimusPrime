
from urllib.parse import urlparse

from .sheetsservice import SheetsService



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

    def dumbificate_prelim_scoresheet(self, quiz_json, roster):
        """This replaces all `importrange(...)` data to a prelim spreadsheet

        Parameters
        ----------
        quiz_json : str
            The draw json element of the scoresheet to dumbificate.

        roster : json
            The roster json to quizzer names from

        """


        file_id = urlparse(quiz_json['url']).path.split("/")[-2]

        value_range_list = []

        # Set Room Number
        values = [[quiz_json['room_num']]]
        value_range_list.append(
            self.generate_value_range_json(
                range = "metadata!B15:B15",
                values = values
            )
        )

        # Set Teams
        for column, team_key in zip("BCD", ["team1", "team2", "team3"]):
            team = quiz_json[team_key]
            values = [[team]]
            for bib in "12345":
                try:
                    quizzer = [q for q in roster if ((q['bib']==bib) and (q['team']==team))][0]
                except IndexError:
                    values.append([""])
                else:
                    values.append([quizzer['moniker']])

            value_range_list.append(
                self.generate_value_range_json(
                    range = "metadata!{0}2:{0}7".format(column),
                    values = values
                )
            )


        self.batch_update_value(
            file_id = file_id,
            value_range_list = value_range_list,
            value_input_option = "USER_ENTERED"
        )

        return self
