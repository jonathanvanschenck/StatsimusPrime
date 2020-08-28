# Documents
## Scoresheet
 - Viewable by: officials
 - Editable by: scorekeepers
 - Imports: quizzer moniker, team moniker, officials names
 - Exports: scoresheet metadata
 - Created with: Quiz number
 - Created at: initialization

## Viewer
 - Viewable by: All
 - Editable by: None
 - Imports: None
 - Exports: quizzer moniker, team moniker
 - Created with: quizzer moniker, team moniker, draw, brackets
 - Created at: initialization

## Statistics
 - Viewable by: None
 - Editable by: Statisticians
 - Imports: scoresheet metadata
 - Exports: None
 - Created With: quizzer moniker, quizzer name, team moniker, draw, bracket
 - Created at: initialization

### Tab: drawlookup
```
 C?, QN,  T1,   T2,   T3,   URL          |   Name,  ABC1, ABC2, ...
 ------------------------------------------------------------------
  x,  1, ABC1, EBC3, NSA1, =`http://...` |      ,    1,      , ...  
   ,  2, ... , ... , ... ,               |      ,     ,     1, ...  
 ...
   ,  A, P_1,  P_4,  P_7 ,               |     ?,    ?,     ?, ...  
   ,  B, P_2,  P_5,  P_8 ,               |     ?,    ?,     ?, ...  
   ,  C, P_3,  P_6,  P_9 ,               |     ?,    ?,     ?, ...  
   ,  D, A_1,  B_1,  C_1 ,               |     ?,    ?,     ?, ...  
   ,  E, A_2,  B_2,  C_2 ,               |     ?,    ?,     ?, ...  
 ...
```
 - Find QN of ith quiz for team index T_ind via: `=IFERROR(INDEX(B3:B,MATCH(i,INDIRECT(CONCATENATE("R2C",T_ind + 7,":C"),0),0),0),"")`
 - Find QN of ith quiz for team T via:
 `=IFERROR(INDEX(B3:B,MATCH(i,INDIRECT(CONCATENATE("R2C", MATCH(T,H1:1,0) + 7,":C"),0),0),0),"")`

### Tab: roster
```
  Bib |         1                                               |    2
 Team |      Name, Moniker, is_rookie, is_captain, is_cocaptain | Name, Moniker, ...
 ---------------------------------------------------------------------------------------------
 ABC1 |  Bob Kyle,  Bob K.,     False,       True,        False | ...
 ABC2 | Joe Smith,  Joe S.,     False,       True,        False | ...
 ...
```

### Tab: schedule
```
          | Room 1                 | Room 2                 | ...
 Time Slot| QN, Team, Score, Place | QN, Team, Score, Place | ...
 --------------------------------------------------------------
  Friday  |   , ABC1, _____, _____ |   ,  ???, _____, _____ | ...
  8:20am  | 1 ,  NSA, _____, _____ | 2 ,  ???, _____, _____ | ...
          |   , EBC2, _____, _____ |   ,  ???, _____, _____ | ...
 --------------------------------------------------------------
 ...
```

### Tab: team_parsed
```   
 Quiz |  1                                 |   2   
 Team | QN, Seat, Score, Placement, Points | ...
 -------------------------------------------------
 ABC1 |  1,   1 ,      ,          ,        | ...
 ABC2 |  5,   2 ,      ,          ,        | ...
 ...
```

### Tab: individual_parsed
```   
                                                             Quiz |  1                                    | 2   
   Name, Team, Bib, is_rookie, is_captain, is_cocaptain, Ave, Acc | QN, C, I, B, BE, 3+PB, CA, CO,  F,  S | ...
 --------------------------------------------------------------------------------------------------------------
    Joe, ABC1,   1,     False,       True,        False,    ,     |  1,  ,  ,  ,    ,    ,    ,  ,   ,    | ...
  Sarah, ABC1,   1,     False,       True,        False,    ,     |  1,  ,  ,  ,    ,    ,    ,  ,   ,    | ...
 ...
```


### Tab: team_raw
```
 Team, QN, Which, Seat, Score, Placement, Points
 ------------------------------------------------
 ABC1,  1,     1,    1,      ,          ,        
 ABC1, 10,     2,    2,      ,          ,      
 ....
```

### Tab: individual_raw
```
 Name, Team, Bib, QN, which, C, I, S, 3PB, 4PB, 5PB, F, CA, CO
 -------------------------------------------------------
  Joe, ABC1,   1,  1,     1,   ,  ,  ,   ,    ,    ,  ,   ,    
  Joe, ABC1,   1, 10,     2,   ,  ,  ,   ,    ,    ,  ,   ,    
  ...
```

### Tab: Utils
 - Table containing bracket weights

### Prep Order
 - [x] Upload Stats doc

 - [x] Update `Utils!C2:C5` with proper weights

 - [x] Inject roster (this creates the team list too)

 - [x] Inject draw (DrawLookup)
 - [x] Calculate number of teams (TN)
 - [x] Calculate total quiz number (TQN)
 - [x] `DrawLookup!L2` fill right (`PASTE_FORMULA`) TN (fill zeros)
 - [x] `DrawLookup!L3` fill right (`PASTE_FORMULA`) TN and down TQN (fill calculate quiz index for team)
 - [x] `DrawLookup!F3` fill down (`PASTE_FORMULA`) TQN (SR lookup cell)
 - [x] `DrawLookup!K3` fill down (`PASTE_FORMULA`) TQN (URL lookup cell)
 - [x] Delete old namedRange `QUIZINDEXLOOKUP` if present
 - [x] Attach NamedRange`QUIZINDEXLOOKUP` as `DrawLookup!L3:`

 - [x] Calculate total number of slots (TSN)
 - [x] `Schedule!A3:F5` copy down (`PASTE_NORMAL`) TSN times (prepare schedule vertical)
 - [x] Calculate number of rooms (TRN)
 - [x] `Schedule!C1:F` copy right (`PASTE_NORMAL`) TRN times (prepare each room)

 - [x] Calculate Semis fill formulas and inject (DrawLookup)

 - [x] `TeamSummary!E2:E` copy down (`PASTE_NORMAL`) TN times
 - [x] Populate post finals ranking ? ? ?

 - [x] Calculated number of prelims (PN)
 - [x] `TeamParsed!E3:I3` fill down (`PASTE_FORMULA`) TN (extract values)
 - [x] `TeamParsed!E1:I` copy (`PASTE_NORMAL`) PN to the right (prep prelims)
 - [x] Generate total team points formula and insert into `TeamParsed!C3`
 - [x] Generate average team points formula and insert into `TeamParsed!D3`
 - [x] `TeamParsed!C3:D3` fill down (`PASTE_FORMULA`) TN (total and average points)
 - [x] `TeamParsed!A3` fill down (`PASTE_FORMULA`) TN (ranking)


 - [x] Inject roster into `IndividualParsed!A3:E`
 - [x] Calculated locations of `SS` columns (SSARRAY =  `{6 + m*12 | for m = 1,...,PN}` and `{6 + (PN + m) * 12 + 1 | for m = 1, 2, 3, 4}`)
 - [x] Insert `=IFERROR(AVERAGE(SSARRAY),"")` into `IndividualParsed!F3`
 - [x] `IndividualParsed!G3:AE3` fill down (`PASTE_FORMULA`) forever (allows for late addition quizzers)
 - [x] Insert 12*(4-1) rows to the right of `AE` (Track all bracket quizzers)
 - [x] `IndividualParsed!T1:AE` copy (`PASTE_NORMAL`) 4-1 times (Add brackets)
 - [x] Insert 12*(PN-1) rows to the right of `R` (Track all prelim quizzes)
 - [x] `IndividualParsed!G1:R` copy (`PASTE_NORMAL`) PN-1 times (Add prelim quizzes)
 - [x] `IndividualParsed!F3` fill down (`PASTE_FORMULA`) forever (calculate averages based on scaled scores)

 - [ ] Generate viewer document

 - [ ] Prepare scoresheets and get URLs
 - [ ] ss_folder share permissions to "everyone with link can view"
 - [ ] Inject URLs (`DrawLookup!F3:F` of the form `="https://google.com"`)


# Main Routines
`sp generate spreadsheets`
 - provide a roster csv: 'name, team_moniker, bib_number, is_rookie, is_captain, is_cocaptain'
 - provide an officials csv 'name, type, room number, email'
 - Either: provide a draw csv OR draw format
 - Either: provide a brackets csv OR brackets format

`sp statistics publish [options]`
 - `bracket` -> update the brackets in `Viewer` every 30 seconds
 - `all` -> dumps everything from `Statistics` into `Viewer`

`sp statistics finalize`
 - call at the end of the meet, to generate quizzer-quiz row tables

# Helper Routines
`sp generate draw`
 - provide number of teams, number of rooms, break location, etc
 - (Optional) provide team list csv

`sp generate brackets`
 - ???

# Extra Routines
`sp update quizzer`
 - Add to rosters
 - Add to IndividualParsed in Stats
`sp update roster`
`sp update officials`
