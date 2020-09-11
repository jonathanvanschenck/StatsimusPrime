# Generate a mock roster for testing

import json
from random import choice

TRN = 3
_TEAMS = ["ABC1","ABC2","ABC3","ABC4","ABC5","ABC6","ABC7","ABC8","ABC9",
          "DEF1","DEF2","DEF3","DEF4","DEF5","DEF6","DEF7","DEF8","DEF9"]
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

team_json = []
id = 0
for team in _TEAMS:
    for i in range(1,choice([4,5,6])):
        name = choice(LETTERS) + choice(LETTERS) + choice(LETTERS) + choice(LETTERS)
        team_json.append({
            "id":"{:0>4}".format(id),
            "team":team,
            "bib":str(i),
            "name":name,
            "moniker":name[:3],
            "is_rookie":"F",
            "is_cap":"FT"[int(i==1)],
            "is_cc":"FT"[int(i==2)]
        })
        id += 1

with open("roster.json","w+") as f:
    f.write(json.dumps(team_json, indent=4))
