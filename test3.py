import statsimusprime.draw as d
p = d.Prelims(nTeams = 18, QpT = 3, nRooms = 2, numblanks = 3).initialize().thermalize(1000)
json = p.generate_json('ABCDEFGHIJKLMNOPQRSTUVWXYZ'[:18],)


# p = d.SemiGraph(d.generate_semi_quizzes("",4)+d.generate_semi_quizzes("1")).assign_links().assign_order()
