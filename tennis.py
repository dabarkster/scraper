from bs4 import BeautifulSoup
import requests
import re
import csv
import os
import datetime
from tinydb import TinyDB, Query
import urllib3
import xlsxwriter
#import openpyxl

#db = TinyDB('./Data/db.json')
#db = TinyDB('db.json')
#db = TinyDB('storage=MemoryStorage')
#db.purge()
#tablePlayers_25_12 = db.table('tablePlayers_25_12')
#tablePlayers_30_14 = db.table('tablePlayers_30_14')

#db.insert({'type': 'apple', 'count': 7})
#print(db.all())

def getPlayerRecord(playerID):
    url = "http://tennislink.usta.com/TeamTennis/Reports/IndividualPlayerRecord.aspx?&MemberID=" + playerID + "&ChampYear=undefined"
    playerRecord = requests.get(url)
    soupRecordResults = BeautifulSoup(playerRecord.content, 'html.parser')
    rows = soupRecordResults.findAll('tr')
    #programs = soupRecordResults.findAll('b')
    tables = soupRecordResults.findAll('table')

    print(tables)

    for row in rows:
        cels = row.findAll('td')
        #print(cels)

    return()


def getFinalScore(matchID):
    url = "http://tennislink.usta.com/TeamTennis/Main/CompletedScoreCard.aspx?MatchTab=M&Preview=Yes&MatchID=" + matchID
    #print(url)
    scoreSummary = requests.get(url)
    soup = BeautifulSoup(scoreSummary.content, 'html.parser')
    table = soup.find('table', id="MatchIndividual1_PageButtonPreview")
    rows = table.findChildren("tr")

    teams = rows[1].findAll('td')
    scores = rows[4].findAll('td')
    
    homeTeam = teams[1].text
    visitTeam = teams[2].text
    homeScore = scores[1].text
    visitScore = scores[2].text
    
    #scores = {'MatchID' : matchID, 'Home' : homeTeam, 'Visitor' : visitTeam, 'HomeScore' : homeScore, 'VisitorScore' : visitScore}
    listScores = [matchID, homeTeam, homeScore, visitScore, visitTeam]
    
    #print(scores)
    return(listScores)



def getResults(matchID, position):
    url = "http://tennislink.usta.com/TeamTennis/Main/CompletedScoreCard.aspx?MatchTab=M&Preview=Yes&MatchID=" + matchID
    print(url)
    matchResults = requests.get(url)
    soup = BeautifulSoup(matchResults.content, 'html.parser')
    table = soup.find('table', {'id': 'MatchIndividual1_tblMatchPreview'})
    rows = table.findChildren("tr")
    datePlayed = rows[2].findAll('font')[3].text.strip()
    matchTop = rows[5].findAll('td')
    homeTeam = matchTop[1].text.strip().replace("*", "")
    visitorTeam = matchTop[2].text.strip().replace("*", "")
    #print(datePlayed)
    #print(homeTeam)
    #print(visitorTeam)

    if   position == "1S" or position == "MS":
        set = '0'
    elif position == "2S" or position == "FS":
        set = '1'
    elif position == "1D" or position == "MD":
        set = '2'
    elif position == "2D" or position == "FD":
        set = '3'
    elif position == "3D" or position == "XD":
        set = '4'

    setID = "MatchIndividual1_Match_Ind_Preview_ctl0" + set + "_tblMatch_Ind_Preview"
    table = soup.find('table', {'id': setID})
    br = soup.find_all("br")
    
    rows = table.findChildren("tr")
    cels = rows[1].findAll('font')
    outcome = cels[0].text.strip()
    if outcome != "Dbl.Default Match":
        match = re.match(r"\s+(.*) is the Winner", cels[0].text)
        winner = match.group(1)
        #print("Cel[1] ##############")
        #print(cels[1].text.strip())
        #print("Cel[2] ##############")
        #print(cels[2].text.strip())
        tempPlayer1 = cels[1].text.replace('\n', '').replace('\r', '')
        playerHome = re.sub(r"\t+","#", tempPlayer1)
        tempPlayer2 = cels[2].text.replace('\n', '').replace('\r', '')
        
        playerVisitor = re.sub(r"\t+","#", tempPlayer2)
		
        if "N/A" in str(cels[1]):
            playerHome1 = "N/A"
            if int(set) > 1:
                playerHome2 = "N/A"
            else:
                playerHome2 = ""
        else:
            playersHome = re.match(r"#([a-zA-Z\s\-\']*)#([a-zA-Z\s\-\']*)?", playerHome)
            playerHome1 = playersHome.group(1)
            if int(set) > 1:
                playerHome2 = playersHome.group(2)
            else:
                playerHome2 = ""

        if "N/A" in str(cels[2]):
            playerVisitor1 = "N/A"
            if int(set) > 1:
                playerVisitor2 = "N/A"
            else:
                playerVisitor2 = ""

        else:
            playersVisitor = re.match(r"#([a-zA-Z\s\-\']*)#([a-zA-Z\s\-\']*)?", playerVisitor)
            playerVisitor1 = playersVisitor.group(1)
            if int(set) > 1:
                playerVisitor2 = playersVisitor.group(2)
            else:
                playerVisitor2 = ""

        score = cels[3].text.split()
        scoreHigh =score[0]
        scoreLow = score[2]
        #print(winner)
        if winner == homeTeam:
            #print('home')
            scoreHome = scoreHigh
            scoreVisitor = scoreLow

        else:
            #print('visitor')
            scoreHome  = scoreLow
            scoreVisitor = scoreHigh

    else:
        playerHome1 = 'N/A'
        playerHome2 = 'N/A'
        playerVisitor1 = 'N/A'
        playerVisitor2 = 'N/A'
        scoreHome = '0'
        scoreVisitor = '0'
        pass

    matchResult = matchID + "," + datePlayed + "," + position + "," + homeTeam + "," + visitorTeam + "," + "," + playerHome1 + "," + playerHome2 + "," + scoreHome + "," + playerVisitor1 + "," + playerVisitor2 + "," + scoreVisitor
    #print(matchResult)
    listResult = (matchID, datePlayed, position, homeTeam, playerHome1, playerHome2, scoreHome, scoreVisitor, visitorTeam, playerVisitor1, playerVisitor2, scoreVisitor)
    
    return(listResult)

def writeExcel(scoreData):
    global worksheet1
    global wsRow
    wsCol = 0
    print(scoreData)
    worksheet1.write_row(wsRow, wsCol, scoreData)
    wsRow += 1


def getMatchSummary(flightID):
    listMatches = []
    matchSummary = requests.get("http://tennislink.usta.com/TeamTennis/Reports/MatchSummary.aspx?Level=F&FlightID=" + flightID + "&ChampYear=undefined")
    soupMatchSummary = BeautifulSoup(matchSummary.content, 'html.parser')
    #print(matchSummary.status_code)
    tables = soupMatchSummary.findChildren('table')
    table = tables[3] #Match summary
    # #print(table)
    rows = table.findChildren("tr")
    #print(len(rows))
    for row in rows:
        #print("rows")
        cels = row.findAll('td')
        link = cels[0].find('a', href=True)

        if link:
            status = cels[9].text
            matchID = link.text
            if status == "Not Played":
                continue
            else:
                listMatches = []
                if "25" in flightID:
                    listMatches.append(matchID)
                    listMatches.append(getResults(matchID, '1S'))
                    listMatches.append(getResults(matchID, '2S'))
                    listMatches.append(getResults(matchID, '1D'))
                    listMatches.append(getResults(matchID, '2D'))
                    listMatches.append(getResults(matchID, '3D'))
                else:
                    listMatches.append(matchID)
                    listMatches.append(getResults(matchID, 'MS'))
                    listMatches.append(getResults(matchID, 'FS'))
                    listMatches.append(getResults(matchID, 'MD'))
                    listMatches.append(getResults(matchID, 'FD'))
                    listMatches.append(getResults(matchID, 'XD'))

                result = getFinalScore(matchID)
                print(listMatches)
                print(result)
                #writeExcel(result)
                #print("\n\n")

        else:
            continue
            
        
    #print(listMatches)
    return(listMatches)


    for cel in cels:
        #print(cel.text)
        link = cel.find('a')
        if link:
            href = link.get('href')
            print(href)
            match = re.match(r"\s+.*:MatchAnchorForMatchBlankSC\((\d+)", href)
            #print(match)
            if match:
                print(match)
                dbKey = match.group(1)
                print(dbKey)
	
        links = row.findChildren('a')
        for link in links:
            href = link.get('href')
            match = re.match(r"\s+.*:MatchAnchorForMatchBlankSC\((\d+)", href)
            if match:
                #print(match)
	            dbKey = match.group(1)

    return()

def getPlayerRoster(flightID):
    print('hi')
    url = "https://tennislink.usta.com/TeamTennis/Reports/PlayerRoster.aspx?Level=F&FlightID=" + flightID + "&ChampYear=undefined"
    
    playerRoster = requests.get(url)
    soupPlayerRoster = BeautifulSoup(playerRoster.content, 'html.parser')
    tables = soupPlayerRoster.findChildren("table")
    table = tables[1]
    rows = table.findChildren("tr")
    links = soupPlayerRoster.find_all('a')
    listAllPlayers = []

    for link in links:
        href = link.get('href')
        match = re.match(r"\s+\w+:(Team|Individual)Anchor\((\d+)", href)
        #print(match)
        if match:
            dbKey = match.group(1)
            if dbKey == "Team":
                teamKey = match.group(2)
                teamName = link.text
            else:
                playerKey = match.group(2)
                playerName = link.text
                #print(teamName)
                #print(name)
                #print("Team: %s" % teamKey)
                #print("Player: %s" % playerKey)
                dictPlayer = {'TeamAnchor' : teamKey, 'TeamName' : teamName, 'IndividualAnchor' : playerKey, 'PlayerName' : playerName}
                #tablePlayers_25_12.insert(dictPlayer)
                #dictPlayer = {playerKey: playerName}
                listAllPlayers.append(dictPlayer)
    print(len(listAllPlayers))
    return(listAllPlayers)

def WriteListToCSV(csv_file, csv_columns, data_list):
    try:
        with open(csv_file, 'a') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', quoting=csv.QUOTE_NONNUMERIC)
            for data in data_list:
                writer.writerow(data)
    except IOError as e:
        errno, strerror = e.args
        print("I/O error({0}): {1}".format(errno, strerror))
    return


def main():
    print('start')
    fileScoreTracker = "Fall2018.xlsm"
    global wsFinalScores
    global wsRow
    #wb = openpyxl.load_workbook(fileScoreTracker)
    #wb.create_sheet('FinalScores')
    #wsFinalScores = wb.get_sheet_by_name('FinalScores')
    #wb.save(fileScoreTracker)
    wsRow = 1
    wsCol = 0
# csv_columns = ['MatchID','Date','Position', 'HomeTeam', 'VisitorTeam', 'HomePlayer1', 'HomePlayer2', 'HomeScore', 'VisitorPlayer1', 'VisitorPlayer2', 'VisitorScore']
# fileResults = r"c:\tmp\results.csv"
# with open(fileResults, 'w') as csvfile:
#     writer = csv.writer(csvfile, dialect='excel', quoting=csv.QUOTE_NONNUMERIC)
#     writer.writerow(csv_columns)
#     WriteListToCSV(fileResults, csv_columns, listMatch)

# csv_columns = ['MatchID','Date','Position', 'HomeTeam', 'VisitorTeam', 'HomePlayer1', 'HomePlayer2', 'HomeScore', 'VisitorPlayer1', 'VisitorPlayer2', 'VisitorScore']
# fileResults = r"c:\tmp\results.csv"
# with open(fileResults, 'w') as csvfile:
#     writer = csv.writer(csvfile, dialect='excel', quoting=csv.QUOTE_NONNUMERIC)
#     writer.writerow(csv_columns)
    #print(getPlayerRoster('154004'))
    flightID_2018f_25_12 = '154004' #2.5(12), 2018 Fall
    flightID_2018f_30_14 = '153997' #3.0(14), 2018 Fall
    #getFinalScore(flightID)
    getMatchSummary(flightID_2018f_30_14)
    #writeExcel()
    #print(getPlayerRoster(flightID_2018f_25_12))
    #print(db.all())
    #workbook.close()
    print('end')
    return()
    
main()

exit()

#matchDict = {matchID + 'x1S': getResults(matchID, '1S')}
# print(matchDict['1S'])
#matchDict[matchID + 'x2S'] = getResults(matchID, '2S')
# print(matchDict['2S'])
#matchDict[matchID + 'x1D'] = getResults(matchID, '1D')
# print(matchDict['1D'])
#matchDict[matchID + 'x2D'] = getResults(matchID, '2D')
# print(matchDict['2D'])
#matchDict[matchID + 'x3D'] = getResults(matchID, '3D')
# print(matchDict['3D'])
#print(matchDict)
#print(getPlayerRoster())
#getPlayerRecord("4623781")
exit(0)


conn_str = (
	    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
	    r'DBQ=D:\Tennis.accdb;'
	    )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
for table_info in crsr.tables(tableType='TABLE'):
	print(table_info.table_name)
