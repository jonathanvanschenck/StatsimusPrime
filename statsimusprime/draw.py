from random import randint,random, choice
from math import ceil, exp, log10

class Quiz:
    """Base class for a single quiz

    The quiz expects up to three teams in the form of single utf-8 characters
    which can be appended onto the quiz by the .push(...) function
    """
    def __init__(self,name,teams = ""):
        self.name = "{: >2}".format(name)
        self.s = teams

    @property
    def s(self):
        return self._s
    @s.setter
    def s(self,string):
        ss = str(string)
        if len(ss) > 3:
            # Ensure that there are only ever three teams / quiz
            raise ValueError("Quiz can only have 3 teams")
        self._s = ss

    @property
    def full(self):
        # True when there are 3 teams in a quiz
        return len(self) == 3

    @property
    def empty(self):
        return len(self) == 0

    def __repr__(self):
        return "<"+self.s+(3-len(self.s))*"_"+">"
    def __str__(self):
        return self.__repr__()

    def __len__(self):
        return len(self.s)

    def push(self,char):
        # Append a team to the quiz (single utf-8 character)
        self.s = self.s + char

    def pop(self,index):
        # Pop a team out of the quiz based of the index of the team
        i = index%3
        v = self.s[i]
        self.s = self.s[:i]+self.s[i+1:]
        return v

class Prelims:
    """Base class to create the draw for prelims

    The program creates an initial guess for the draw and then tries to optimize
    according to several metrics:
        - Minimize the number of back-to-back quizzes for each teams
        - Minimize the number of triple-back-to-back (hat tricks) for each team
        - Minimize the number of times a team repeated quizzes another team
        - Minimize the number of times a team repeatedly quizzes in the same room
        - Prevent a team for quizzing in two places at the same time
    The class does this optimization using a modification of the Metropolis-Hastings
    algorithm, where each of the above events are given an "energy", higher for
    events which are less desireable (e.g. coincident quizzing and hat tricks).
    The total energy of the draw is calculated as the of all the event energies.
    The draw then is "thermalized" by repeatedly trying to interchange two teams
    (or two whole quizzes), calculating the new total energy after the interchange.
    The interchange is "accepted" probablistically based on the two total energies
    using:
                             / exp(-(Enew-Eold)/kT) if Eold < Enew
        Prob of Acceptance =|
                             \  1                     otherwise
    where "kT" is a specified "thermal energy" which exists in the sytem (Here,
    you can think of this as an average amount of undesirable event which is
    tolerated in the system). Ideally, you should iterate through this algorithm
    (implemented as .thermalize(...)) twice: once with a nonzero kT value, which
    will randomize the draw, and then again with a very small kT value to "freeze"
    the draw into an energetic minimum. See the example below.

    Implemented also is a Simulated Annealing algorithm, which is basically just
    MH, but you start with a high temperature and then slowly drop it each iteration,
    this is accessable through the ".anneal(...)" method. Many thanks to Ted Towers
    (github.com/tcubed) for the idea to add this functionality!

    :: Example Usage MH Algorithm::
    >>> T = 18 # number of teams
    >>> QpT = 7 # number of quizzes per team
    >>> R = 6 # number of rooms
    >>> breaklocation = 0.5 # location of the day break
    >>> d = Prelims(T,QpT,R,breaklocation)
    >>> d.initialize() # Create an initial (bad) draw
    >>> print("Thermalizing") # Randomize
    >>> d.thermalize(10**4, kT = 0.1, alpha = 0.2, verbose=True)
    >>> print("Freezing") # Find optimal
    >>> d.thermalize(10**4, kT = 0.001, alpha = 0.2, verbose=True)
    >>> stats = d.get_stats(verbose=True) # Print out statistics on the draw
    >>> print(d.to_text()) # Format into text

    :: Example Usage SA Algorithm::
    >>> T = 18 # number of teams
    >>> QpT = 7 # number of quizzes per team
    >>> R = 6 # number of rooms
    >>> breaklocation = 0.5 # location of the day break
    >>> d = Prelims(T,QpT,R,breaklocation)
    >>> d.initialize() # Create an initial (bad) draw
    >>> print("Annealing") # Randomize
    >>> d.anneal(10**5, verbose=True)
    >>> stats = d.get_stats(verbose=True) # Print out statistics on the draw
    >>> print(d.to_text()) # Format into text
    """

    # Energy penalties for each "undesirable" event type
    back_to_back = 0.3
    hat_trick = 1.0
    already_seen = 0.1
    already_quizzed = 0.05
    currently_quizzing = 10.0

    def __init__(self, nTeams, QpT, nRooms, breakloc = None, numblanks = None):
        """
        nTeams : The number of teams in the Meet
        QpT : The number of quizzes per team
        nRoom : The number of rooms
        brealoc : A float specifying the approximate location of the day break
        numblanks : The number of blank quizzes to include
        """

        assert nTeams/3 >= nRooms, "Too many rooms!"
        assert QpT%3==0 or nTeams%3==0, "Either teams or qpT must be divisible by 3"

        S = ceil(nTeams*QpT/3/nRooms)
        self.B = numblanks or S*nRooms-QpT*nTeams//3
        self.R = nRooms
        self.T = nTeams
        self.Q = nTeams*QpT//3
        self.S = ceil((self.Q+self.B) / self.R)
        self.E = 0.0
        bl = breakloc or 1.1
        self.breakindex = int(round(self.S*bl))
        self.qpt = QpT
        self.draw = []
        i = 0
        for s in range(self.S):
            self.draw.append([])
            for _ in range(min(self.Q+self.B-i,self.R)):
                self.draw[-1].append(Quiz(str(i+1)))
                i += 1
        self.Tlist = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"[:self.T]
        self.Tquiz = {t:[] for t in self.Tlist}

    def __repr__(self):
        l = [", ".join([q.name+str(q) for q in s]) for s in self.draw]
        l.insert(self.breakindex,"")
        return "\n".join(l)

    def _generate_open(self,char):
        # Generate all the not filled quizzes
        slots_to_fill = ceil((self.Q) / self.R)
        for si in range(slots_to_fill):
            if si == slots_to_fill-1:
                rooms_to_fill = self.Q - (slots_to_fill - 1) * self.R
            else:
                rooms_to_fill = self.R
            for ri in range(rooms_to_fill):
                if not self.draw[si][ri].full:# and not char in q.s:
                    yield (si,ri)

    def get_quiz_energy(self,char,si,ri):
        """Calaculates the 'energy' of a team being inserted into a quiz.

        Note, this function assumes the team `char` is NOT already in quiz
        'self.draw[si][ri]'. If it IS, you will need to call self.pop(char,si,ri)
        then this function, the self.push(char,si,ri).
        """

        # Check for back-to-back
        try:
            btb = any(char in q.s for q in self.draw[si-1])
        except IndexError:
            # Happens for quizzes in the first slot
            btb = False
        if si == self.breakindex:
            # Override for quizzes immediately after the break
            btb = False

        # Check for hat-tricks
        try:
            ht = btb and any(char in q.s for q in self.draw[si-2])
        except IndexError:
            # Happens for quizzes in the first or second slot
            ht = False
        if si == self.breakindex:
            # Override for quizzes immediately after the break
            ht = False

        # Check if a team is already quizzing
        cq = any(char in q.s for _ri,q in enumerate(self.draw[si]))

        # Check the number of other times `char` is already quizzing the other
        #  teams in the quiz
        seen = 0
        for other_char in self.draw[si][ri].s:
            for _si,_ri in self.Tquiz[other_char]:
                if not (_si == si and _ri == ri) and char in self.draw[_si][_ri].s:
                    seen += 1

        # Check the number of other times `char` is quizzing in this room
        quizzed = 0
        for _si,_ri in self.Tquiz[char]:
            quizzed += int(ri==_ri)

        return cq*self.currently_quizzing\
                + quizzed*self.already_quizzed\
                + btb*self.back_to_back\
                + ht*self.hat_trick\
                + seen*self.already_seen

    def get_total_energy(self):
        """Calculates the total energy of the draw
        """
        E = 0.0
        for si,s in enumerate(self.draw):
            for ri,q in enumerate(s):
                for char in q.s:
                    self.pop(char,si,ri)
                    E += self.get_quiz_energy(char,si,ri)
                    self.push(char,si,ri)
        return E

    def pop(self,char,si,ri):
        """Remove a team `char` from the quiz: self.draw[ri][si]
        """
        try:
            val = self.draw[si][ri].pop(self.draw[si][ri].s.index(char))
        except Exception as E:
            raise E
        else:
            self.Tquiz[char].pop(self.Tquiz[char].index((si,ri)))

        return val

    def push(self,char,si,ri):
        """Add a team `char` to the quiz: self.draw[ri][si]
        """
        try:
            self.draw[si][ri].push(char)
        except Exception as E:
            raise E
        else:
            self.Tquiz[char].append((si,ri))

    def initialize(self):
        """Create an initial draw before thermalization

        This function tries to create an rough guess at a good draw by iterating
        through each team and trying to place it in the lowest energy spot available.
        For a serial method, I think this is the best one can do.

        Note, this initial draw is deterministic (it will always turn out the
        same way).

        Note, this initial draw does not gaurentee that a team won't be scheduled for
        two quizzes at the same time, but this will only happen if the draw parameters
        are 'tight' (meaning there isn't a lot of play: lots of rooms or not very
        many blank quizzes).
        """
        for _ in range(self.qpt):
            for char in self.Tlist:
                E = [(si,ri,self.get_quiz_energy(char, si, ri)) for si,ri in self._generate_open(char)]
                si,ri,E = sorted(E,key=lambda e: e[2])[0]
                self.push(char,si,ri)
        self.E = self.get_total_energy()
        return self

    def interchange_team(self,char1,si1,ri1,char2,si2,ri2,kT = 1.0):
        """Attempt to interchange two teams using the Metropolis Algorithm
        """
        # Interchange two teams
        self.pop(char1,si1,ri1)
        self.pop(char2,si2,ri2)
        self.push(char2, si1, ri1)
        self.push(char1, si2, ri2)
        # Get the new total energy
        Enew = self.get_total_energy()
        deltaE = Enew-self.E
        try:
            prob = exp(-(deltaE)/kT)
        except (ZeroDivisionError, OverflowError):
            prob = 0.0
        if deltaE < 0:
            # Accept the interchange
            self.E += deltaE
            return True, deltaE
        elif random() < prob:
            # Accept the interchange
            self.E += deltaE
            return True, deltaE
        else:
            # Reject the interchange
            self.pop(char2, si1, ri1)
            self.pop(char1, si2, ri2)
            self.push(char1, si1, ri1)
            self.push(char2, si2, ri2)
            return False, deltaE

    def interchange_quiz(self,qn1,qn2,kT = 1.0):
        """Attempt to interchange two quizzes using the Metropolis Algorithm
        """
        override = (kT == 0.0)
        # Get the slot index and room index from the quiz number
        si1,ri1 = qn1 // self.R, qn1 % self.R
        si2,ri2 = qn2 // self.R, qn2 % self.R

        # Interchange the two quizzes
        char1 = []
        for char in self.draw[si1][ri1].s:
            char1.append(self.pop(char,si1,ri1))
        char2 = []
        for char in self.draw[si2][ri2].s:
            char2.append(self.pop(char,si2,ri2))

        for char in char1:
            self.push(char,si2,ri2)
        for char in char2:
            self.push(char,si1,ri1)

        # Get the new total energy
        Enew = self.get_total_energy()
        deltaE = Enew-self.E
        try:
            prob = exp(-(deltaE)/kT)
        except (ZeroDivisionError, OverflowError):
            prob = 0.0
        if deltaE < 0:
            # Accept the interchange
            self.E += deltaE
            return True, deltaE
        elif random() < prob:
            # Accept the interchange
            self.E += deltaE
            return True, deltaE
        else:
            # Reject the interchange
            for char in char1:
                self.pop(char,si2,ri2)
            for char in char2:
                self.pop(char,si1,ri1)
            for char in char1:
                self.push(char,si1,ri1)
            for char in char2:
                self.push(char,si2,ri2)
            return False, deltaE

    def _thermalization_step(self,kT = 1.0, alpha = 0.1):
        """Iterate once through the Metropolis Algorithm
        """
        if random() > alpha:
            # Try a team interchange
            ci1,ci2 = randint(0,self.T-1),randint(1,self.T-1)
            ci2 = (ci1 + ci2) % self.T
            char1,char2 = self.Tlist[ci1],self.Tlist[ci2]
            qi1,qi2 = randint(0,self.qpt-1), randint(0,self.qpt-1)
            si1,ri1 = self.Tquiz[char1][qi1]
            si2,ri2 = self.Tquiz[char2][qi2]
            return self.interchange_team(char1,si1,ri1,char2,si2,ri2,kT=kT)
        else:
            # Try a quiz interchange
            qn1,qn2 = randint(0,self.Q+self.B-1),randint(0,self.Q+self.B-1)
            return self.interchange_quiz(qn1,qn2,kT=kT)

    def thermalize(self, N, kT = 0.5, alpha = 0.1,verbose=False):
        """Run through the Metropolis Algorithm to randomize the draw
        """
        j = 0
        for i in range(N):
            self._thermalization_step(kT=kT,alpha=alpha)
            if (i % (N//20)) == 0:
                j += 1
                if verbose:
                    print("{: >3}% : E = {:.1f} : E/T = {:.2f} : E/Q = {:.3f}".format(5*j,self.E,self.E/self.T,self.E/(3*self.Q)))
        return self

    def anneal(self, N, kTmax = 5, kTmin = 1e-3, alpha = 0.2, verbose = False, log = True):
        """Run through the Simulated Annealing algorithm to randomize the draw
        """
        if log:
            lkTmax,lkTmin = log10(kTmax), log10(kTmin)
            step = (lkTmin-lkTmax)/N
            kT_list = [10.0**(lkTmax + step*i) for i in range(N)]
        else:
            step = (kTmin-kTmax)/N
            kT_list = [kTmax + step*i for i in range(N)]
        j = 0
        for i,kT in enumerate(kT_list):
            self._thermalization_step(kT,alpha)
            if (i % (N//20)) == 0:
                j += 1
                if verbose:
                    print("{: >3}% : kT = {: >1.3f} : E = {:.1f} : E/T = {:.2f} : E/Q = {:.3f}".format(5*j, kT,self.E,self.E/self.T,self.E/(3*self.Q)))
        return self

    def get_stats(self,verbose=False):
        """Generate statistics on the draw
        """
        stats = {}

        for char in self.Tlist:
            stat = {}

            # Get the number of times `char` will quiz in each room
            for ri in range(self.R):
                stat[ri] = 0
            for si,ri in self.Tquiz[char]:
                stat[ri] += 1

            # Get the number of times `char` will quiz each other tea
            for other_char in self.Tlist:
                stat[other_char] = 0
                for si,ri in self.Tquiz[char]:
                    stat[other_char] += int(other_char in self.draw[si][ri].s)

            # Check for coincident quizzing and back-to-back quizzes
            stat['cq'] = [False,[]]
            stat['btb'] = [0,[]]
            for i in range(self.qpt-1):
                si1,ri1 = self.Tquiz[char][i]
                for j in range(i+1,self.qpt):
                    si2,ri2 = self.Tquiz[char][j]
                    if si1 == si2 and ri1 == ri2:
                        stat['cq'][0] = True
                        stat['cq'][1].append((si1,ri1,si2,ri2))
                    if ((si1 == si2 + 1) or (si1 == si2 - 1)) and si2 != self.breakindex:
                        stat['btb'][0] += 1
                        stat['btb'][1].append((si1,ri1,si2,ri2))

            # Check for hat tricks
            stat['ht'] = [0,[]]
            l = len(stat['btb'][1])
            for i in range(max(0,l-1)):
                for j in range(i+1,l):
                    si1,ri1,si2,ri2 = stat['btb'][1][i]
                    si3,ri3,si4,ri4 = stat['btb'][1][j]
                    if max(si1,si2,si3,si4) == min(si1,si2,si3,si4) + 2:
                        stat['ht'][0] += 1
                        stat['ht'][1].append((
                            min(si1,si2,si3,si4),
                            int(0.5*((si1+si2+si3+si4)-max(si1,si2,si3,si4)-min(si1,si2,si3,si4))),
                            max(si1,si2,si3,si4)
                        ))


            stats[char] = stat

        if verbose:
            # Quizzes are valid if no team is schedule to quiz in two places at
            #  the same time
            valid = True
            for k,v in stats.items():
                if v['cq'][0]:
                    valid = False
                    print("Conflict at:",k,v['cq'])
            print("Strictly Valid?",valid)

            print("# of Hat Tricks:",len({k:v['ht'][1] for k,v in stats.items() if v['ht'][0] > 0}))

            print("Back-to-backs:")
            for k,v in {k:v['btb'][1] for k,v in stats.items() if v['btb'][0] > 0}.items():
                print("\t",k,":",len(v))

            print("Room Assignments:")
            print("\t  "," ".join([str(i+1) for i in range(self.R)]))
            for char in self.Tlist:
                print("\t",char," ".join([str(stats[char][i]) for i in range(self.R)]))


            print("Team Corelation:")
            print("   "," ".join(self.Tlist))
            for char in self.Tlist:
                print(char+" |"," ".join([str(stats[char][other_char]) for other_char in self.Tlist]))


            print("Final Energy:")
            EE = {
                "E  ": self.E,
                "E/T": self.E / self.T,
                "E/Q": self.E / (3 * self.Q)
            }
            for k,v in EE.items():
                print("\t",k,"=","%.3f" % v)

            print("Draw:")
            print(self)

        return stats

    def to_text(self,swaplist = None):
        """Render the draw to a nice text form

        Teams are comma delimited, quizzes are semi-colon delimited, and slots
        are cariage-return delimited
        """
        l = [
            ";".join([
                ",".join([
                    char for char in str(q)[str(q).index("<")+1:str(q).index(">")]
                ]) for q in s ]) for s in self.draw
        ]
        l.insert(self.breakindex,"")
        return "\n".join(l)

    @classmethod
    def from_text(cls,file_path):
        """Reconstruct a draw from a file genereated by `.to_text()`
        """
        with open(file_path) as f:
            draw = []
            for l in f:
                draw.append([q.split(",") for q in l.strip().split(";")])

        team = []
        for s in draw:
            for q in s:
                for char in q:
                    team.append(char)

        swaplist = [char for char in set(team) if char != '' and char != '_']

        nTeams = len(swaplist)
        QpT = len([char for char in team if char == 'A'])
        breakless = [s for s in draw if s != [['']]]
        nRooms = len(breakless[0])
        try:
            breakloc = draw.index([['']]) / len(breakless)
        except ValueError:
            breakloc = 1.1
        numblanks = len([char for char in team if char == "_"])//3
        new_cls = cls(nTeams,QpT,nRooms,breakloc,numblanks)

        for si,s in enumerate(breakless):
            for ri,q in enumerate(s):
                if q != ['_','_','_']:
                    for char in q:
                        new_cls.push(char,si,ri)

        new_cls.E = new_cls.get_total_energy()

        return new_cls

    def generate_json(self, team_list):
        """Create a json object representing the draw

        Parameters
        ----------
        team_list : list
            List of the names of each team, must match the length of the number
            of teams provided at instantiation, or bad things will happen.

        returns : obj = [
            ...,
            {
                quiz_num: "2", # 1 indexed
                slot_num: "1", # 1 indexed
                room_num: "2", # 1 indexed
                team1: "",     # Team name
                team2: "",     # Team name
                team3: "",     # Team name
                type: "P"      # Prelim type quiz
            },
            ...
        ]
        """
        swaplist = {char:team for char,team in zip(self.Tlist,team_list)}
        swaplist[""] = ""
        quizzes = []
        qi = 0
        for si, s in enumerate(self.draw):
            for ri, q in enumerate(s):
                if not q.empty:
                    qi += 1
                    quiz = {
                        "quiz_num": str(qi),
                        "slot_num": str(si+1),
                        "room_num": str(ri+1),
                        "type": "P"
                    }
                    for i in range(3):
                        quiz["team{}".format(i+1)] = swaplist[q.s[i]]
                    quizzes.append(quiz)

        return quizzes


class RoundRobin(Prelims):
    def __init__(self, nTeams, QpT = 3):#, nSlots, extraSlots):
        nRooms = 1

        # numblanks = (nRooms * nSlots) - (nTeams * QpT // 3 - extraSlots)
        Prelims.__init__(self, nTeams, QpT, nRooms)#, numblanks = numblanks)

    def generate_json(self, room_num, offset_bracket, offset_slot):
        """Generate the json representation

        Parameters
        ----------
        room_num : int or string
            The (1 indexed) number of the room where the quizzes will take place

        offset_bracket : int (1 or 2)
            An integer representing the offset for the bracket: either 1 for A or
            2 for B.

        offset_slot : int
            The slot number of the last prelim.
        """
        team_list = ["P_{}".format(9*offset_bracket+1+i) for i in range(len(self.Tlist))]
        swaplist = {char:team for char,team in zip(self.Tlist,team_list)}
        swaplist[""] = ""
        quizzes = []
        qi = 0
        for si, s in enumerate(self.draw):
            for ri, q in enumerate(s):
                if not q.empty:
                    qi += 1
                    quiz = {
                        "quiz_num": "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[qi-1] + str(offset_bracket),
                        "slot_num": str(offset_slot+si+1),
                        "room_num": str(room_num),
                        "type": "ABC"[offset_bracket-1]
                    }
                    for i in range(3):
                        quiz["team{}".format(i+1)] = swaplist[q.s[i]]
                    quizzes.append(quiz)

        return quizzes



quiz_to_teams = {
    "A": {"team1":"P_1","team2":"P_4","team3":"P_7"},
    "B": {"team1":"P_2","team2":"P_5","team3":"P_8"},
    "C": {"team1":"P_3","team2":"P_6","team3":"P_9"},
    "D": {"team1":"A_1","team2":"B_1","team3":"C_1"},
    "E": {"team1":"A_2","team2":"B_2","team3":"B_2"},
    "F": {"team1":"A_3","team2":"B_3","team3":"C_3"},
    "G": {"team1":"D_2","team2":"D_3","team3":"E_1"},
    "H": {"team1":"E_2","team2":"E_3","team3":"F_1"},
    "I": {"team1":"G_2","team2":"G_3","team3":"H_1"},
    "J": {"team1":"D_1","team2":"G_1","team3":"I_1"},
    "K": {"team1":"P_1","team2":"P_2","team3":"P_3"}
}
def get_teams(quiz_name, bracket_offset):
    """Get the json for teams

    Parameters
    ----------
    quiz_name : str
        The name of the quiz: A, B, ...

    bracket_offset : int
        The offset for the bracket: 0 for semis, 1 for A, 2 for B
    """
    teams = quiz_to_teams[quiz_name].copy()
    if bracket_offset > 0:
        for k in teams.keys():
            if teams[k].split("_")[0].lower() != "p":
                teams[k] = teams[k].split("_")[0]+str(bracket_offset)+"_"+teams[k].split("_")[1]
            else:
                teams[k] = teams[k].split("_")[0]+"_"+str(int(teams[k].split("_")[1])+9*bracket_offset)
    return teams


bracket_types = {
    "full" : {
             #slot offset, room_num
        "A": [1,1],
        "B": [1,2],
        "C": [1,3],
        "D": [2,1],
        "E": [2,2],
        "F": [2,3],
        "G": [3,1],
        "H": [3,2],
        "I": [4,1],
        "J": [5,1],
    },
    "condensed" : {
        "A": [1,1],
        "B": [1,2],
        "C": [2,1],
        "D": [3,1],
        "E": [3,2],
        "F": [4,1],
        "G": [5,1],
        "H": [5,2],
        "I": [6,1],
        "J": [7,1],
    },
    "condensed_left" : {
        "A": [1,1],
        "B": [1,2],
        "C": [2,1],
        "D": [3,1],
        "E": [3,2],
        "F": [4,1],
        "G": [5,1],
        "H": [6,1],
        "I": [7,1],
        "J": [8,1],
    },
    "condensed_right" : {
        "A": [1,3],
        "B": [2,2],
        "C": [2,3],
        "D": [3,3],
        "E": [4,2],
        "F": [4,3],
        "G": [5,2],
        "H": [5,3],
        "I": [6,2],
        "J": [7,2],
    },
    "none" : {},
    "finals_only" : {
        "K": [1,1]
    }
}

def generate_bracket_json(bracket_style, slot_offset, num_brackets = 1, finals_repeats = [1,1,1]):
    quizzes = []
    if bracket_style == "full":
        bracket = bracket_types['full']
        for bracket_offset in range(num_brackets):
            finals_repeat = finals_repeats[bracket_offset]
            for i in range(9+int(finals_repeat==1)):
                quiz_name = "ABCDEFGHIJ"[i]
                si,ri = bracket[quiz_name]
                quiz = {
                    "quiz_num": quiz_name + ["","1","2"][bracket_offset],
                    "slot_num": str(slot_offset+si),
                    "room_num": str(ri + 3*bracket_offset),
                    "type": "SABC"[bracket_offset]
                }
                quiz.update(get_teams(quiz_name, bracket_offset))
                quizzes.append(quiz)
            if finals_repeat > 1:
                for i in range(finals_repeat):
                        quiz_name = "J"
                        si,ri = bracket[quiz_name]
                        quiz = {
                            "quiz_num": quiz_name + ["","1","2"][bracket_offset] + "({})".format(i+1),
                            "slot_num": str(slot_offset+si+i),
                            "room_num": str(ri + 3*bracket_offset),
                            "type": "SABC"[bracket_offset]
                        }
                        quiz.update(get_teams(quiz_name, bracket_offset))
                        quizzes.append(quiz)
    elif bracket_style == "condensed":
        for bracket_offset in range(num_brackets):
            bt = ['condensed_left','condensed_right','condensed'][bracket_offset%2 + int(((bracket_offset + 1) == num_brackets) and (num_brackets%2 == 1))]
            bracket = bracket_types[bt]
            finals_repeat = finals_repeats[bracket_offset]
            for i in range(9+int(finals_repeat==1)):
                quiz_name = "ABCDEFGHIJ"[i]
                si,ri = bracket[quiz_name]
                quiz = {
                    "quiz_num": quiz_name + ["","1","2"][bracket_offset],
                    "slot_num": str(slot_offset+si),
                    "room_num": str(ri + 3*(bracket_offset//2)),
                    "type": "SABC"[bracket_offset]
                }
                quiz.update(get_teams(quiz_name, bracket_offset))
                quizzes.append(quiz)
            if finals_repeat > 1:
                for i in range(finals_repeat):
                        quiz_name = "J"
                        si,ri = bracket[quiz_name]
                        quiz = {
                            "quiz_num": quiz_name + ["","1","2"][bracket_offset] + "({})".format(i+1),
                            "slot_num": str(slot_offset+si+i),
                            "room_num": str(ri + 3*(bracket_offset//2)),
                            "type": "SABC"[bracket_offset]
                        }
                        quiz.update(get_teams(quiz_name, bracket_offset))
                        quizzes.append(quiz)
    if bracket_style == "finals_only":
        bracket = bracket_types['finals_only']
        for bracket_offset in range(num_brackets):
            finals_repeat = finals_repeats[bracket_offset]
            for i in range(int(finals_repeat==1)):
                quiz_name = "K"
                si,ri = bracket[quiz_name]
                quiz = {
                    "quiz_num": quiz_name + ["","1","2"][bracket_offset],
                    "slot_num": str(slot_offset+si),
                    "room_num": str(ri + 3*bracket_offset),
                    "type": "SABC"[bracket_offset]
                }
                quiz.update(get_teams(quiz_name, bracket_offset))
                quizzes.append(quiz)
            if finals_repeat > 1:
                for i in range(finals_repeat):
                        quiz_name = "K"
                        si,ri = bracket[quiz_name]
                        quiz = {
                            "quiz_num": quiz_name + ["","1","2"][bracket_offset] + "({})".format(i+1),
                            "slot_num": str(slot_offset+si+i),
                            "room_num": str(ri + 3*bracket_offset),
                            "type": "SABC"[bracket_offset]
                        }
                        quiz.update(get_teams(quiz_name, bracket_offset))
                        quizzes.append(quiz)
    return quizzes

def generate_semis_json(nTeams, slot_offset, finals_repeats = [1,1,1], bracket_style = 'condensed'):
    quizzes = generate_bracket_json(
        bracket_style = bracket_style,
        slot_offset = slot_offset,
        num_brackets = nTeams//9,
        finals_repeats = finals_repeats
    )
    NT = nTeams % 9
    room_num = max([int(q['room_num']) for q in quizzes])+1
    max_slot = max([0]+[int(q['slot_num']) for q in quizzes if not "J" in q['quiz_num']])+1
    QpT = 3#*((max_slot - slot_offset) // NT)
    if NT >= 3:
        quizzes += RoundRobin(
            nTeams = NT,
            QpT = QpT
        ).initialize().anneal(1000).generate_json(
            room_num,
            offset_bracket = nTeams//9,
            offset_slot = slot_offset
        )
    return quizzes


if __name__ == "__main__":
    T = 18 # number of teams
    QpT = 7 # number of quizzes per team
    R = 6 # number of rooms
    breaklocation = 0.5 # location of the day break
    b = Prelims(T,QpT,R,breaklocation)
    b.initialize()
    print("Thermalizing")
    b.thermalize(10**4, kT = 0.1, alpha = 0.2, verbose=True)
    print("Freezing")
    b.thermalize(10**4, kT = 0.001, alpha = 0.2, verbose=True)
    stats = b.get_stats(verbose=True)
    print(b.to_text())
