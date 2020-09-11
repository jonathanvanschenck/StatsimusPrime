# StatsimusPrime
Automated System for generating Quiz Meet Stats

# Operation
## Initial Set Up
Get a `credentials.json` file: [see here](https://developers.google.com/docs/api/quickstart/python). Then dump it into whatever folder you are going to run `StatsimusPrime` prime in.

Set up a virtual python3 environment and activate it:
```bash
 $ python3 -m venv venv
 $ source venv/bin/activate
 (venv) $ pip install -r requirements.txt
```

Currently, `StatsimusPrime` is run through a `python3` shell, which can be configured using:
```bash
 (venv) $ python3 test-tools/start_shell.py
 >>> # run everything from here
```

## Setting up a meet
### Preliminary Work
Delete the `.env` file from the folder where you are running `StatsimusPrime`. Note, this won't permanently destroy anything if you have been using `m.push_env()` to back up your environment file online, it only allows you to start over with a new meet. Technically, you don't need to do this step, because later the `.env` file would get overridden anyway, but it is a good practice to always start fresh. After you delete the `.env` file, be sure to restart the `StatsimusPrime` shell.

Next, create a google drive folder where you want all the meet documents to get uploaded. Then grab the share link (or just id) of that folder, and give it to `StatsimusPrime` using:
```python3
 >>> m.initialize_env('your_folder_share_url_here') # This many take some time to run . . .
```
Note, if you give it the url for a folder which has already been partially set up with a quiz meet, this method should auto-detect that and try to reconstruct the meet in memory as best as possible.

### Loading a roster and a draw
Now you need to create the `roster` and `draw` objects. There are a few helper tools for this. For instance, if you are part of PNW, the `test-tools/pnw_create_roster.py` script will convert the `registration_list.csv` downloaded from https://pnwquizzing.org/tool/registration_list into the proper json file: `roster.json`. Once you have a roster file, you can load it into `StatsimusPrime` using:
```python3
 >>> m.load_roster('roster.json') # load the roster into memory
 >>> m.save_env() # save the roster onto your hard drive backup environment file (`.env`)
 >>> m.push_env() # duplicate a copy of environment in google drive, for backup if you lose your `.env` file
```

If you have previously loaded a roster into `StatsimusPrime` using `m.load_roster('roster.json')`, there is a basic method for generating a draw. It takes a number of parameters, which you can read about in the source code: `statsimusprime/manager.py`:
```python3
 >>> m.generate_draw_from_roster({ ... your paramters here ... }, verbose = True) # creates `draw.json` file
 >>> m.load_draw('draw.json') # load the draw into memory
 >>> m.save_env() # save the draw onto your hard drive backup environment file (`.env`)
 >>> m.push_env() # duplicate a copy of environment in google drive, for backup if you lose your `.env` file
```
Alternatively, if you created the `draw.json` file some other way, or modified it after generation, just run the last three lines above.

### Generating the meet
Once you have the roster and draw loaded into `StatsimusPrime`, you can then generate the full quiz meet:
```python3
 >>> m.generate_quiz_meet() # This will definitely take some time
```

After the meet has finished generating, you can then give edit privileges to the `scoresheets` folder in google drive to all the scorekeepers and officials. Note, DON'T give edit or view privileges to the main folder in google drive, 'cause then people can screw up your main statistics google sheet. Just don't do it, it is a security risk. Also DO NOT change the share link privileges on the `scoresheets` folder. The `scoresheets` folder already has the permissions set to "Anyone With Link Can View" and if you change that, everything will break. Rather, use the email feature to directly give edit privileges to each of your officials individually.

What you CAN do is give the share link (which is "view" only) to the `scoresheets` folder to any spectator that you want. You can click on "Get Shareable Link" and post that somewhere public for parents/coaches/quizzers to find. Just don't change what the "link" access grants to people.

### Running a Meet
During Prelims (and also brackets), you can open up the main `Statistics` google sheet in the top folder on google drive. Everything should be automated, except approving quizzes as "complete". To do this, you can double check the scoresheet in the `scoresheets` folder, and when you are satisfied, you can mark it complete by putting a capital "Y" in column A of tab `DrawLookup` in the `Statistics` google sheet. Everything else should be handled.

Once brackets start, you will need to tell `StatsimusPrime` to periodically update the static statistics viewer: `scoresheets/_Statistics_Viewer` in drive. If you don't do this, then bracket scoresheets won't know which teams are supposed to participate. This isn't the end of the world, because the scorekeeper can just manually put them in, but it is a nice touch. Also, the quizzers won't necessarily know which quiz they are supposed to be in next, becuase they only have access to the stats viewer. To set this up, run
```python3
 >>> m.update_brackets_every(seconds = 60) # this runs until you hit ctrl+c
```

### Finishing a Meet
Currently, `StatsimusPrime` cannot calculate final team placement, so you will have to do that. You can stick them in `TeamSummary` columns H and I. Look out for version 1.0 when that will get changed.

At the end of the meet, you can publish (copy into `scoresheets/_Statistics_Viewer`) all the team and individual statistics by running:
```python3
 >>> m.publish_quiz_meet()
```

Eventually, I will fix the download options, which will let you download a static copy of everything, and also make the individual and team statistics more database-friendly. For now, too bad.

# To Do (version 1.0)
 - [x] Add draw creation support
 - [x] Add bracket update support
 - [ ] Add post-finals team ranking support (for brackets)
 - [ ] fix "copy over" no blanks problem
 - [x] Fix the readme
 - [ ] fix static backup support
 - [ ] clean up the documentation

# To Do (version 2.0)
 - [ ] Add permissions support (officials)
 - [ ] Add cli support
 - [ ] Add meet-to-meet YTD support
 - [ ] switch to logger rather than print statements
 - [ ] Add pre-finals team ranking support (for RoundRobin finals)
 - [ ] add RoundRobin finals support
 - [ ] Add post-finals team ranking support (for RoundRobin)
