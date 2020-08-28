import json
from random import choice

TRN = 3
_TEAMS = ["ABC1","ABC2","ABC3","ABC4","ABC5","ABC6","ABC7","ABC8","ABC9",
          "DEF1","DEF2","DEF3","DEF4","DEF5","DEF6","DEF7","DEF8","DEF9"]
LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
draw = []
for _ in range(3):
    TEAMS = 1*_TEAMS
    while len(TEAMS) > 0:
        draw.append([]) # add slot
        for __ in range(TRN):
            q = []
            for ___ in range(3):
                q.append(choice(TEAMS))
                TEAMS.pop(TEAMS.index(q[-1]))
            draw[-1].append(q)

draw_json = []
qn = 1
for si,s in enumerate(draw):
    for ri,q in enumerate(s):
        draw_json.append({
            "quiz_num":str(qn),
            "slot_num":str(si+1),
            "room_num":str(ri+1),
            "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
            "team1":q[0],
            "team2":q[1],
            "team3":q[2],
            "url":"",
            "type":"P"
        })
        qn += 1

si += 1
draw_json.append({
    "quiz_num":"A",
    "slot_num":str(si+1),
    "room_num":str(1),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"P_1",
    "team2":"P_4",
    "team3":"P_7",
    "url":"",
    "type":"S"
})
draw_json.append({
    "quiz_num":"B",
    "slot_num":str(si+1),
    "room_num":str(2),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"P_2",
    "team2":"P_5",
    "team3":"P_8",
    "url":"",
    "type":"S"
})
draw_json.append({
    "quiz_num":"C",
    "slot_num":str(si+1),
    "room_num":str(3),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"P_3",
    "team2":"P_6",
    "team3":"P_9",
    "url":"",
    "type":"S"
})
si += 1
draw_json.append({
    "quiz_num":"D",
    "slot_num":str(si+1),
    "room_num":str(1),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"A_1",
    "team2":"B_1",
    "team3":"C_1",
    "url":"",
    "type":"S"
})
draw_json.append({
    "quiz_num":"E",
    "slot_num":str(si+1),
    "room_num":str(2),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"A_2",
    "team2":"B_2",
    "team3":"C_2",
    "url":"",
    "type":"S"
})
draw_json.append({
    "quiz_num":"F",
    "slot_num":str(si+1),
    "room_num":str(3),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"A_3",
    "team2":"B_3",
    "team3":"C_3",
    "url":"",
    "type":"S"
})
si += 1
draw_json.append({
    "quiz_num":"G",
    "slot_num":str(si+1),
    "room_num":str(1),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"D_2",
    "team2":"D_3",
    "team3":"E_1",
    "url":"",
    "type":"S"
})
draw_json.append({
    "quiz_num":"H",
    "slot_num":str(si+1),
    "room_num":str(2),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"E_2",
    "team2":"E_3",
    "team3":"F_1",
    "url":"",
    "type":"S"
})
si += 1
draw_json.append({
    "quiz_num":"I",
    "slot_num":str(si+1),
    "room_num":str(1),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"G_2",
    "team2":"G_3",
    "team3":"H_1",
    "url":"",
    "type":"S"
})
si += 1
draw_json.append({
    "quiz_num":"J",
    "slot_num":str(si+1),
    "room_num":str(1),
    "slot_time":"Friday {}:{:0>2}".format(8+si//3,20*(si%3)),
    "team1":"D_1",
    "team2":"G_1",
    "team3":"I_1",
    "url":"",
    "type":"S"
})

with open("draw.json","w+") as f:
    f.write(json.dumps(draw_json, indent=4))

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
