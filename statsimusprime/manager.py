# Adapted from https://developers.google.com/docs/api/quickstart/python

import pickle
import os
import json
from time import sleep


from urllib.parse import urlparse
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from apiclient.http import MediaFileUpload

# from statsimusprime.service import DriveService, StatsService, ScoresheetService

from statsimusprime.service.driveservice import DriveService
from statsimusprime.service.statsservice import StatsService
from statsimusprime.service.scoresheetservice import ScoresheetService
from statsimusprime.draw import Prelims, generate_semis_json

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
        self.viewerfp = os.path.join(wd,'templates','Viewer_template.xlsx')
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
        """Initialize the quizzing environment

        This method will first check if an env file exists in the top folder. if
        if does, it will assume that the folder is already setup correctly and
        load that environment. Else, it will try to find all the relevent files
        inside the top folder to build up the environment file from scratch. Any
        files it doesn't find will be uploaded.

        top_folder_id_or_url : str
            The url or id of the top folder to be used for a particular meet.
        """
        top_folder_id = urlparse(top_folder_id_or_url).path.split("/")[-1]

        self.env = {
            'top_folder_id': top_folder_id,
            'env_id': None,
            'trash_id': None,
            'scoresheets_id': None,
            'stats_id': None,
            'ss_template_id': None,
            'viewer_id': None,
            'roster': [],
            'draw': [],
            'officials': [],
            "bracket_weights": {"P":1.0,"S":1.0,"A":0.7,"B":0.5}
        }


        # Create the environment file if it doesn't exist (or truncate if it does)
        with open(self.envfp,"w") as f:
            f.write("{}")

        # Check if any of the files or folders already exist
        files = self.drive_service.get_all_children(top_folder_id)
        for _f in files:
            if _f['mimeType'] == 'application/json':
                if _f['name'] == 'env':
                    # If the environment is found, download it and load!
                    print("Found a environment document in Google Drive")
                    self.env["env_id"] = _f['id']
                    return self.pull_env().load_env()

            elif _f['mimeType'] == 'application/vnd.google-apps.folder':
                if _f['name'] == 'trash':
                    self.env['trash_id'] = _f['id']
                    print("Found a trash folder in Google Drive")
                elif _f['name'] == 'scoresheets':
                    self.env['scoresheets_id'] = _f['id']
                    print("Found a scoresheets folder in Google Drive")
            elif _f['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                if _f['name'] == 'Statistics':
                    self.env['stats_id'] = _f['id']
                    print("Found a stats document in Google Drive")
                elif _f['name'] == 'Scoresheet_template':
                    self.env['ss_template_id'] = _f['id']
                    print("Found a scoresheet template in Google Drive")
                # elif _f['name'] == '_Statistics_Viewer':
                #     self.env['viewer_id'] = _f['id']
                #     print("Found a viewer in Google Drive")

        # If any don't, create them
        if self.env['trash_id'] is None:
            self.env['trash_id'] = self.drive_service.create_folder(
                name='trash',
                parent_folder_id = top_folder_id
            ).get('id')
        if self.env['scoresheets_id'] is None:
            # Create the non-existent scoresheets folder
            self.env['scoresheets_id'] = self.drive_service.create_folder(
                name='scoresheets',
                parent_folder_id = top_folder_id
            ).get('id')
            # Add viewer priviledges
            self.drive_service.publish_file(self.env['scoresheets_id'])
        else:
            # Else, crawl the existent scoresheet folder to look for the stats viewer
            files = self.drive_service.get_all_children(self.env['scoresheets_id'])
            for _f in files:
                if _f['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                    if _f['name'] == '_Statistics_Viewer':
                        self.env['viewer_id'] = _f['id']
                        print("Found a viewer in Google Drive")

        if self.env['stats_id'] is None:
            self.env['stats_id'] = self.drive_service.upload_excel_as_sheet(
                name = 'Statistics',
                file_path = self.statsfp,
                parent_folder_id = top_folder_id
            ).get('id')

        if self.env['viewer_id'] is None:
            self.env['viewer_id'] = self.drive_service.upload_excel_as_sheet(
                name = '_Statistics_Viewer',
                file_path = self.viewerfp,
                parent_folder_id = self.env['scoresheets_id']
            ).get('id')

        if self.env['ss_template_id'] is None:
            self.env['ss_template_id'] = self.drive_service.upload_excel_as_sheet(
                name = 'Scoresheet_template',
                file_path = self.ssfp,
                parent_folder_id = top_folder_id
            ).get('id')

        if self.env['env_id'] is None:
            # Initialize the drive side of the env file
            self.env['env_id'] = self.drive_service.upload_json(
                name = "env",
                file_path = self.envfp,
                parent_folder_id = top_folder_id
            ).get('id')


        # Create the backup, reload env, then push the backup onto drive
        return self.save_env().load_env().push_env()


    def load_env(self):
        """Loads an `.env` file from the working directory into the Manager

        This includes activating all the gsuite services, and doing some helpful
        preliminary calculations about the quiz meet (like calculating the number
        of teams, etc.)

        This method is either called during instantiation, or by `.initialize_env()`
        if this is the first time the manager is being run for a particular meet
        """
        with open(self.envfp) as f:
            self.env = json.load(f)

        # Activate drive and sheets services
        self.drive_service.id = self.env['top_folder_id']
        self.drive_service.trash_id = self.env['trash_id']

        self.stats_service.id = self.env['stats_id']
        self.stats_service.viewer_id = self.env['viewer_id']
        self.stats_service.retrieve_meet_parameters(
            self.env['roster'],
            self.env['draw']
        )

        self.ss_service.id = self.env['ss_template_id']

        return self

    def save_env(self):
        """This saves the current environment as `.env` in the current working directory
        """
        with open(self.envfp, "w") as f:
            json.dump(self.env, f, indent=4)

        return self

    def push_env(self):
        """This pushes the local `.env` file on the cloud
        """
        self.drive_service.update_json(
            file_id = self.env['env_id'],
            file_path = self.envfp
        )
        return self

    def pull_env(self):
        """This pulls down the `env` file from the cloud in the current working directory

        NOTE: this method only modifies the `.env` file saved locally on your
        harddrive, and does NOT update the `.env` object which is referenced by
        the manager for most other methods. You must call `.load_env()` after
        this method if you want to the manager to update its internal environment
        variable -- which you probably do if you are planning to use the Manager
        for anything else.
        """
        self.drive_service.download_json(
            file_id = self.env['env_id'],
            destination_file_path = self.envfp
        )
        return self

    def load_roster(self, file_path):
        """Loads a roster json file into the manager env object

        Note, this method does not update either the local `.env` file, or
        the `env` file in the cloud. Must call `.save_env().push_env()` to do so.

        Further, if you have previously called `.generate_quiz_meet()` and have
        since changed the roster (say, by adding a new quizzer) you will also need
        to reconfigure the roster in the stats and viewer documents by calling
        `.push_roster()`

        JSON format should be a list of quizzer json objects:
            {
                "id": "0024",         # must be unique and consistent between meets
                "team": "ABC1",
                "bib": "1",           # must be 1-5
                "name": "John Doe",   # Private value, not publish publically
                "moniker": "John D.", # Published publically
                "is_rookie": "F",     # must be "T"/"F"
                "is_cap": "T",        # must be "T"/"F"
                "is_cc": "F"          # must be "T"/"F"
            }

        Parameters
        ----------
        file_path : str
            Path the the roster json file
        """
        with open(file_path) as f:
            self.env['roster'] = json.load(f)

        return self

    def generate_draw_from_roster(self, draw_params = {}, verbose = True):
        """Generate a draw + bracket + finals json from the loaded roster

        Note, this method assumes that self.env['roster'] already exists and has
        all the correct teams loaded into it.

        Note, this method does NOT load the generated draw into self.env['draw'],
        it only saves a files called `draw.json` in the current working directory
        which can then be loaded via `.load_draw('draw.json')` and then have the
        environment saved and pushed with those methods.

        Parameters
        ----------
        draw_params : dict : default = {}
            A dictionary of parameters for generating the draw. Everything has
            a default value, so providing `{}` will simply create the default draw.
            The dictionary takes the following items:
            {
                skip_round_robin : boolean : default = True
                    If skip_round_robin, then only full brackets will be generated,
                    and any teams scoring below the lowest bracket after prelims
                    will simply be done quizzing. If !skip_round_robin AND the
                    nTeams%9>=3, then the lowest ranked teams will be given a
                    round robin-style bracket.

                finals_repeat : list of integers : default = [1, 1, 1]
                    This list specifies the number of times finals could repeat in
                    each bracket. For instance, if there are as many as 4 finals in
                    top-9, but only a single finals quiz in con A and con B, then
                    [4,1,1].

                bracket_style : str : default = 'condensed'
                    The type of bracket to use for finals. There are two valid
                    options. "full" specifies that all brackets will require 3
                    rooms, and the entire semifinals will only use 4 slots. This
                    is really only useful if you arn't running a consolation bracket,
                    or happen to have 6/9 rooms -- unlikely. "condensed" specifies
                    that each bracket will use 2 rooms (though, they can share
                    the second room with another bracket, so 2 brackets can fit
                    in only 3 rooms total). "finals_only" only add a "J" quiz

                num_rooms : int : default = 4
                    The number of rooms available at the quiz meet.

                QpT : int (multiple of 3) : default = 6
                    The number of prelim quizzes that each team will have. Unless
                    the number of teams at the meet happens to be a multiple of 3,
                    QpT MUST be a multiple of 3, otherwise the algorithm that
                    creates the draw may fail. You have been warned.

                num_blanks : integer or None : default = None
                    This is the number of blank ('break') quizzes to include in prelims.
                    If None is given, then the correct number of blanks to exactly
                    fill up a prelim draw rectangle will be used.

            }

        verbose : boolean
            If verbose, all the internal methods will give full printouts of their
            results
        """

        finals_repeats = draw_params.pop("finals_repeats", [1,1,1])
        bracket_style = draw_params.pop("bracket_style", 'condensed')
        assert bracket_style in ['full', 'condensed', 'finals_only'], "Only valid bracket styles are: `full` and `condensed`"
        num_rooms = draw_params.pop("num_rooms", 4)
        QpT = draw_params.pop("QpT", 6)
        num_blanks = draw_params.pop("num_blanks", None)
        skip_round_robin = draw_params.pop("skip_round_robin", True)

        # Document me!
        annealing_steps = draw_params.pop("annealing_steps", 10**4)
        slots_on_friday = draw_params.pop("slots_on_friday", 8)
        slots_before_lunch = draw_params.pop("slots_on_friday", 3)
        friday_start_time = draw_params.pop("friday_start_time", [6+12,20])
        saturday_start_time = draw_params.pop("saturday_start_time", [9,0])
        lunch_break_time = draw_params.pop("lunch_break_time", [1,0])
        prelim_semi_break_time = draw_params.pop("lunch_break_time", [0,20])
        minutes_per_quiz = draw_params.pop("minutes_per_quiz", 20)

        if len(draw_params.keys()) > 1:
            print("Unused parameters: ",", ".join(draw_params.keys()))

        team_list = list(set([q['team'] for q in self.env['roster']]))
        nTeams = len(team_list)

        if verbose:
            print("Generating Prelims, this may take a few minutes . . . ")
        prelim = Prelims(
            nTeams = nTeams,
            QpT = QpT,
            nRooms = num_rooms,
            numblanks = num_blanks
        ).initialize().anneal(annealing_steps, verbose = verbose)
        if verbose:
            prelim.get_stats(verbose=True)

        quizzes = prelim.generate_json(team_list)
        slot_offset = max([int(q['slot_num']) for q in quizzes])

        if verbose:
            print("Generating Brackets quiz . . . ")

        if skip_round_robin:
            if verbose:
                print("Skipping Round Robin for bottom teams . . .")
            nTeams -= nTeams%9

        quizzes += generate_semis_json(
            nTeams = nTeams,
            slot_offset = slot_offset,
            finals_repeats = finals_repeats,
            bracket_style = bracket_style
        )

        if verbose:
            print("Attaching time slots to quizzes . . .")

        for quiz in quizzes:
            quiz['url'] = ''

            # Parse the slot time
            slot = int(quiz['slot_num'])
            if slot <= slots_on_friday:
                template = "Fri {}:{:0>2} {}"
                h,m = friday_start_time
                m += minutes_per_quiz * (slot - 1)
            else:
                template = "Sat {}:{:0>2} {}"
                h,m = saturday_start_time
                adj_slot = slot - slots_on_friday
                m += minutes_per_quiz * (adj_slot - 1)
                if adj_slot > slots_before_lunch:
                    h += lunch_break_time[0]
                    m += lunch_break_time[1]
                if quiz['type'] != "P":
                    h += prelim_semi_break_time[0]
                    m += prelim_semi_break_time[1]
            h += m//60
            m = m%60
            AMPM = ["AM","PM"][h>=12]
            h = 1 + ((h-1)%12)
            quiz['slot_time'] = template.format(h,m,AMPM)


        if verbose:
            print("Saving to `draw.json` . . ")
        with open('draw.json', "w+") as f:
            json.dump(quizzes, f, indent = 4)

        if verbose:
            print("Printing Draw . . . ")
            num_slots = max([int(q['slot_num']) for q in quizzes])
            num_rooms = max([int(q['room_num']) for q in quizzes])
            draw = [[(17+int(r==0)*(12+3))*" " for r in range(num_rooms)] for s in range(num_slots)]
            for q in quizzes:
                if q['room_num'] == "1":
                    ll = "{: >12} | {: >2};{: >4},{: >4},{: >4}".format(
                        q['slot_time'],
                        q['quiz_num'],
                        q['team1'],
                        q['team2'],
                        q['team3']
                    )
                else:
                    ll = "{: >2};{: >4},{: >4},{: >4}".format(
                        q['quiz_num'],
                        q['team1'],
                        q['team2'],
                        q['team3']
                    )
                draw[int(q['slot_num'])-1][int(q['room_num'])-1] = ll
            print(" | ".join([12*" "]+["{: ^17}".format("Room %d" % (r+1)) for r in range(num_rooms)]))
            for s in draw:
                print(" | ".join(s))
        return self

    def load_draw(self, file_path):
        """Loads a draw json file into the manager env object

        Note, this method does not update either the local `.env` file, or
        the `env` file in the cloud. Must call `.save_env().push_env()` to do so.

        JSON format should be a list of quiz json objects:
            {
                "quiz_num": "1",            # The quiz number/name
                "slot_num": "1",            # Slot number (must be integer >= 1)
                "room_num": "1",            # Room number (must be integer >= 1)
                "slot_time": "Friday 8:00", # Project time of quiz
                "team1": "ABC3",            # See note below
                "team2": "DEF1",            # See note below
                "team3": "DEF5",            # See note below
                "url": "",                  # Leave blank
                "type": "P"                 # Specifies type of quiz, must be
                                            #  P(relim),S(emi finals),(con) A,(con) B
            }

        Note: the `team1` of the quiz json object must take be one of three forms:
            1) If the quiz is a prelim quiz, the team must be a valid team identifier
                which correspondes to one of the teams in the `roster` object.

            2) If the quiz is a Semi finals/Consolations quiz, it CAN take the
                form: `P_i` where i is an integer from 1 to the total number of
                teams. This indicates that which ever team places in ith place
                after prelims will quiz in this place: e.g. `P_3` indicates that
                the 3rd placed team after prelims will quiz here.

            3) If the quiz is a Semi finals/Consolations quiz, it CAN take the
                form: `{quiz_num}_i` where `{quiz_num}` is the number/name of a
                previous quiz (provided the quiz isn't called `P` becuase that
                will break everything, just rename it) and `i` is 1, 2 or 3
                indicating a placement in that quiz. For example: `A_3` means
                that whoever took 3rd in quiz A will quiz here. Or `B3_1` means
                whoever took first in quiz B3 will quiz here.

        Parameters
        ----------
        file_path : str
            Path the the draw json file
        """
        with open(file_path) as f:
            self.env['draw'] = json.load(f)

        return self

    def download_static_image(self,fp=None):
        _fp = fp or os.getcwd()

        # Create a temporary backup folder
        bu_id = self.drive_service.create_folder(
            name = "backup",
            parent_folder_id = self.env['top_folder_id']
        ).get('id')

        print("Copying Statistics")
        # Create a stats backup
        bu_stats_id = self.drive_service.copy_to(
            file_id = self.env['stats_id'],
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
        for ss in self.drive_service.get_all_children(self.env['scoresheets_id']):
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

    def generate_scoresheets(self, verbose = True):
        """Makes a copy of the ss_template for each quiz in the environment
        """

        # Remove everything from the scoresheets folder which is not the viewer
        for file in self.drive_service.get_all_children(self.env['scoresheets_id']):
            if file['id'] != self.env['viewer_id']:
                self.drive_service.move_to_trash(file['id'])

        # Generate all the scoresheets and save their urls into the environment
        for i in range(len(self.env['draw'])):
            quiz_num = self.env['draw'][i]['quiz_num']
            if verbose:
                print('Generating quiz {}'.format(quiz_num))

            response = self.drive_service.copy_to(
                file_id = self.env['ss_template_id'],
                name = quiz_num,
                destination_folder_id = self.env['scoresheets_id'],
                fields = "id, webViewLink"
            )

            self.ss_service.set_quiz_number_for(
                file_id = response.get('id'),
                quiz_num = quiz_num
            )

            self.env['draw'][i]['url'] = response.get('webViewLink')


        # Add urls into stats document
        self.stats_service.update_ss_urls(self.env['draw'])

        # Backup the enviroment
        return self.save_env().push_env()

    def generate_quiz_meet(self):
        # Step 1: Prepare stats document
        self.stats_service.retrieve_meet_parameters(
            self.env['roster'],
            self.env['draw']
        ).set_bracket_weights(self.env['bracket_weights'])\
         .set_roster(self.env['roster'])\
         .set_draw(self.env['draw'])\
         .initialize_schedule()\
         .initialize_team_summary()\
         .set_team_parsed()\
         .set_individual_parsed(self.env['roster'])

        # Step 2: copy over viewer document
        self.stats_service.initialize_viewer()\
            .copy_over_draw()\
            .copy_over_roster()#\
            # .copy_over_team_summary()\
            # .copy_over_individual_summary()

        # Step 3: modify the stats template to import data correctly
        self.ss_service.initialize_global_variables(
            self.drive_service.get_file_url(self.env['viewer_id'])
        )

        # Step 4: generate the score sheets
        self.generate_scoresheets()

        return self

    def push_roster(self):
        """Republishes the roster in the stats and viewer documents

        Note, this should only be called after `.generate_quiz_meet()` IF the
        roster has been changed using `.load_roster(...)` i.e. becuase of a late
        added or droped quizzer.
        """
        self.stats_service.retrieve_meet_parameters(
            self.env['roster'],
            self.env['draw']
        ).set_roster(self.env['roster'])\
         .set_individual_parsed(self.env['roster'])\
         .copy_over_roster()
        return self

    def update_brackets_every(self, seconds = 60, ds = 5):
        print("Updating Brackets every {} seconds".format(seconds))
        try:
            while True:
                for i in range(seconds//ds):
                    print("\rNext Update in: {} seconds         ".format(seconds-i*ds),end="")
                    sleep(ds)
                self.stats_service.copy_over_draw()
                print("\rUpdated"+20*" ")
        except KeyboardInterrupt:
            print("\rEnding Updates"+20*" ")

        return self

    def publish_quiz_meet(self):
        self.stats_service.copy_over_team_summary()\
            .copy_over_individual_summary()

        return self

    def dumbificate_prelims(self):
        """This removes all `importranges` from prelim scoresheets

        This is a backup measure if prelim scoresheets are failing to import properly.
        """
        for quiz in self.env['draw']:
            if quiz['type'] == "P":
                self.ss_service.dumbificate_prelim_scoresheet(
                    quiz_json = quiz,
                    roster = self.env['roster']
                )

        return self
