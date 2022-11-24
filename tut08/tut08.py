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
