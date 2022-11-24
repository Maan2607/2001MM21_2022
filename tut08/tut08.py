def scorecard():
	pass
#importing required libraries
import openpyxl  
import pandas as pd
import os

from datetime import datetime
start_time = datetime.now()

#Reading files of innings of Pakistan and India
Indian_inn = open("india_inns2.txt","r+") 
Pakista_inn = open("Pak_inns1.txt","r+") 
teams = open("teams.txt","r+")
Team_squads = teams.readlines()

pak_team = Team_squads[0]
pak_players = pak_team[23:-1:].split(",")

India_team = Team_squads[2]
Ind_players = India_team[20:-1:].split(",")


lst_ind=Indian_inn.readlines() 
for i in lst_ind:
    if i=='\n':
        lst_ind.remove(i)
      

lst_pak=Pakista_inn.readlines() 
for i in lst_pak:
    if i=='\n':
        lst_pak.remove(i)

wb = openpyxl.Workbook()
sheet = wb.active

# batting [runs,ball,4s,6s,sr]
# bowling [over,medan,runs,Wickets, NB, WD, ECO]
ind_fow=0                         #indian fall of wickets
pak_fow=0                         # pakistan fall of wickets
out_pb={}                         # Out batter of pakistan
ind_bowlers={}
ind_batter={}

pak_bats={}
pak_bowlers={}
pak_byes=0
pak_bt=0             #total runs given by pak bowlers
for l in lst_pak:
    x=l.index(".")
    over_pak=l[0:x+2]
    temp=l[x+2::].split(",")
    running_ball=temp[0].split("to") #0 2

    if f"{running_ball[0].strip()}" not in ind_bowlers.keys() :
        ind_bowlers[f"{running_ball[0].strip()}"]=[1,0,0,0,0,0,0]   #[over0,medan1,runs2,Wickets3, NB4, WD5, ECO6]
    elif "wide" in temp[1]:
        pass
    elif "byes" in temp[1]:                        # run count for byes
        if "FOUR" in temp[2]:
            pak_byes+=4
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
        elif "1 run" in temp[2]:
            pak_byes+=1
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
        elif "2 runs" in temp[2]:
            pak_byes+=2
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
        elif "3 runs" in temp[2]:
            pak_byes+=3
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
        elif "4 runs" in temp[2]:
            pak_byes+=4
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
        elif "5 runs" in temp[2]:
            pak_byes+=5
            ind_bowlers[f"{running_ball[0].strip()}"][0]+=1

    else:
        ind_bowlers[f"{running_ball[0].strip()}"][0]+=1
    
    if f"{running_ball[1].strip()}" not in pak_bats.keys() and temp[1]!="wide":
        pak_bats[f"{running_ball[1].strip()}"]=[0,1,0,0,0] #[runs,ball,4s,6s,sr]
    elif "wide" in temp[1] :
        pass
    else:
        pak_bats[f"{running_ball[1].strip()}"][1]+=1
    

    if "out" in temp[1]:                                      #scoring for a out 
        ind_bowlers[f"{running_ball[0].strip()}"][3]+=1
        if "Bowled" in temp[1].split("!!")[0]:                           #if bowled
            out_pb[f"{running_ball[1].strip()}"]=("b" + running_ball[0])
        elif "Caught" in temp[1].split("!!")[0]:                         #if got caught
            w=(temp[1].split("!!")[0]).split("by")
            out_pb[f"{running_ball[1].strip()}"]=("c" + w[1] +" b " + running_ball[0])
        elif "Lbw" in temp[1].split("!!")[0]:                             # got out by lbw
            out_pb[f"{running_ball[1].strip()}"]=("lbw  b "+running_ball[0])

    

    if "no run" in temp[1] or "out" in temp[1] :                     # scoring for no run on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=0
        pak_bats[f"{running_ball[1].strip()}"][0]+=0
    elif "1 run" in temp[1]:                                        # scoring for 1 run on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=1
        pak_bats[f"{running_ball[1].strip()}"][0]+=1
    elif "2 runs" in temp[1]:                                       # scoring for 2 run on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=2
        pak_bats[f"{running_ball[1].strip()}"][0]+=2
    elif "3 runs" in temp[1]:                                       # scoring for 3 run on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=3
        pak_bats[f"{running_ball[1].strip()}"][0]+=3
    elif "4 runs" in temp[1]:                                       # scoring for 4 run on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=4
        pak_bats[f"{running_ball[1].strip()}"][0]+=4
    elif "FOUR" in temp[1]:                                        # scoring for Four  on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=4
        pak_bats[f"{running_ball[1].strip()}"][0]+=4
        pak_bats[f"{running_ball[1].strip()}"][2]+=1
    elif "SIX" in temp[1]:                                         # scoring for six  on current ball
        ind_bowlers[f"{running_ball[0].strip()}"][2]+=6
        pak_bats[f"{running_ball[1].strip()}"][0]+=6
        pak_bats[f"{running_ball[1].strip()}"][3]+=1
    elif "wide" in temp[1]:                                         # scoring for wide  on current ball
        if "wides" in temp[1]:
            # print(temp[1][1])
            ind_bowlers[f"{running_ball[0].strip()}"][2]+=int(temp[1][1])
            ind_bowlers[f"{running_ball[0].strip()}"][5]+=int(temp[1][1])
        else:
            ind_bowlers[f"{running_ball[0].strip()}"][2]+=1
            ind_bowlers[f"{running_ball[0].strip()}"][5]+=1

