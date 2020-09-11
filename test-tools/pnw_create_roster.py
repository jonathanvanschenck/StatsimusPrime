import json
from sys import argv

try:
    fp_source = argv[1]
except IndexError:
    fp_source = "registration_list.csv"

try:
    fp_dest = argv[1]
except IndexError:
    fp_dest = "roster.json"

with open(fp_source) as f:
    keys = f.readline().strip().split(",")
    roster_list = [{k:v.strip('"') for k,v in zip(keys,l.strip().split(","))} for l in f]

quizzers = []
for i,q in enumerate(roster_list):
    if q['Role'].lower() == 'quizzer':
        try:
            is_rookie = q['Rookie'].strip().lower()[0]=="y"
        except:
            is_rookie = False
        quizzers.append({
            "id":"{:0>4}".format(i),
            "team":"".join(q['Team'].strip().split(" ")),
            "bib":q['Bib'].strip(),
            "name":q['Name'].strip(),
            "moniker":q['Name'].strip().split(" ")[0]+" "+q['Name'].strip().split(" ")[-1][0]+".",
            "is_rookie":"FT"[is_rookie],
            "is_cap":"FT"[q['Captain'].strip().lower()=="c"],
            "is_cc":"FT"[q['Captain'].strip().lower()=="cc"]
        })

for quizzer in quizzers:
    if quizzer['team'] == '':
        print("Quizzer: {} ({}) has no team".format(quizzer['name'],quizzer['id']))
    if not(quizzer['bib'] in '12345'):
        print("Quizzer: {} ({}) has invalid bib".format(quizzer['name'],quizzer['id']))


with open(fp_dest,"w+") as f:
    f.write(json.dumps(quizzers, indent=4))
