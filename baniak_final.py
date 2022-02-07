from mimetypes import init
from openpyxl import load_workbook
from openpyxl.descriptors.base import String
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import colorsys
import random
import baniak_network
import baniak_helper
from plotly.subplots import make_subplots

workBook = load_workbook('/Users/przemyslawszafulski/Developer/BaniakBaniaka/Chronologia_BCU.xlsx')
workSheet = workBook["Chronologia BCU"]

def _get_colors(num_colors):
    colors=[]
    for i in np.arange(0., 360., 360. / num_colors):
        hue = i/360.
        lightness = (50 + np.random.rand() * 10)/100.
        saturation = (90 + np.random.rand() * 10)/100.
        color = colorsys.hls_to_rgb(hue, lightness, saturation)
        hexString = "#" + "".join("%02X" % int(round(i*255)) for i in color)
        colors.append(hexString)
    return colors

class Person:
    def __init__(self, searchName, displayName, groupName, x = 200, y = 200):
        self.searchName = searchName
        self.displayName = displayName
        self.groupName = groupName
        self.x = x
        self.y = y
    
    searchName = ""
    displayName = ""
    groupName = ""
    sessionCount = 0
    sessionLengthInSeconds = 0
    x = None
    y = None

    def color(self):
        groupName = self.groupName
        if groupName == "HOBO":
            return "#F5AB39"
        elif groupName == "SUM":
            return "#4ECDA9"
        elif groupName == "KŁM":
            return "#DD85FF"
        elif groupName == "GON":
            return "#1E34DB"
        elif groupName == "Popiół":
            return "#FF0019"
        elif groupName == "Inne":
            return "#FAF500"
        else:
            return "#555555"
    
    def sessionLengthString(self):
        total = self.sessionLengthInSeconds
        seconds = total % 60
        total -= seconds
        minutes = (total / 60) % 60
        total -= minutes * 60
        hours = (total / 3600)
        string = ""
        if (hours < 10):
            string += "0" + str(int(hours))
        else:
            string += str(int(hours))
        string += ":"
        if (minutes < 10):
            string += "0" + str(int(minutes))
        else:
            string += str(int(minutes))
        string += ":"
        if (seconds < 10):
            string += "0" + str(int(seconds))
        else:
            string += str(int(seconds))
        return string
    
    
    def sessionLengthInHours(self):
        return self.sessionLengthInSeconds / 3600

class Session:
    people = []
    name = ""
    teamName = ""
    sessionLengthInSeconds = 0


    def color(self):
        groupName = self.teamName
        if groupName == "HOBO":
            return "#F5AB39"
        elif groupName == "SUM":
            return "#4ECDA9"
        elif groupName == "KŁM":
            return "#DD85FF"
        elif groupName == "GON":
            return "#1E34DB"
        elif groupName == "Popiół":
            return "#FF0019"
        elif groupName == "Inne":
            return "#FAF500"
        else:
            return "#555555"

    def sessionLengthString(self):
        total = self.sessionLengthInSeconds
        seconds = total % 60
        total -= seconds
        minutes = (total / 60) % 60
        total -= minutes * 60
        hours = (total / 3600)
        string = ""
        if (hours < 10):
            string += "0" + str(int(hours))
        else:
            string += str(int(hours))
        string += ":"
        if (minutes < 10):
            string += "0" + str(int(minutes))
        else:
            string += str(int(minutes))
        string += ":"
        if (seconds < 10):
            string += "0" + str(int(seconds))
        else:
            string += str(int(seconds))
        return string
    
    def sessionLengthInHours(self):
        return self.sessionLengthInSeconds / 3600


def extractSessionLengthInSeconds(row):
    timeColumn = workSheet["G"]
    i = 1
    while True:
        value = timeColumn[row - i].value
        if (value != None):
            return (value.hour * 60 + value.minute) * 60 + value.second
        else:
            i += 1
    
def extractSessionName(row):
    sessionNameColumn = workSheet["A"]
    i = 1
    while True:
        value = sessionNameColumn[row - i].value
        if (value != None):
            return value
        else:
            i += 1

def extractTeamName(row):
    teamNameColumn = workSheet["C"]
    i = 1
    while True:
        value = teamNameColumn[row - i].value
        if (value != None):
            return value
        else:
            i += 1

hexagons = [[1, 0], [0.5, 0.86], [-0.5, 0.86], [-1, 0], [-0.5, -0.86], [0.5, -0.86]]
multiplier = 100

pentagon = [[0.80, 0.26], [0, 0.85], [-0.8, 0.26], [-0.5, -0.68], [0.5, -0.68]]
pentagonMultiplier = 100

bigMultiplier = 500

offsetHoboX = pentagon[2][0] * bigMultiplier
offsetHoboY = pentagon[2][1] * bigMultiplier

offsetKlmX = pentagon[1][0] * bigMultiplier
offsetKlmY = pentagon[1][1] * bigMultiplier

offsetSumX = pentagon[0][0] * bigMultiplier
offsetSumY = pentagon[0][1] * bigMultiplier

offsetGonX = pentagon[3][0] * bigMultiplier
offsetGonY = pentagon[3][1] * bigMultiplier

offsetPopX = pentagon[4][0] * bigMultiplier
offsetPopY = pentagon[4][1] * bigMultiplier


# offsetHoboX = 0
# offsetHoboY = 0

# offsetKlmX = -300
# offsetKlmY = -300

# offsetSumX = 300
# offsetSumY = -300

# offsetGonX = 300
# offsetGonY = 300

# offsetPopX = -300
# offsetPopY = 300

people = [Person("Nikos", "Nikos", "HOBO", hexagons[0][0] * multiplier + offsetHoboX, hexagons[0][1] * multiplier + offsetHoboY), 
          Person("Dzieran", "Dzieran", "HOBO", hexagons[3][0] * multiplier + offsetHoboX, hexagons[3][1] * multiplier + offsetHoboY), 
          Person("Karol", "Karol", "HOBO", hexagons[2][0] * multiplier + offsetHoboX, hexagons[2][1] * multiplier + offsetHoboY), 
          Person("Lewy", "Lewy", "HOBO", hexagons[1][0] * multiplier + offsetHoboX, hexagons[1][1] * multiplier + offsetHoboY), 
          Person("Kozi", "Kozi", "HOBO", hexagons[4][0] * multiplier + offsetHoboX, hexagons[4][1] * multiplier + offsetHoboY), 
          Person("Słowik", "Słowik", "HOBO", hexagons[5][0] * multiplier + offsetHoboX, hexagons[5][1] * multiplier + offsetHoboY), 
          Person("Tremiszewski", "Wojciech Tremiszewski", "KŁM", pentagon[0][0] * pentagonMultiplier + offsetKlmX, pentagon[0][1] * pentagonMultiplier + offsetKlmY), 
          Person("Giza", "Abelard Giza", "KŁM", pentagon[2][0] * pentagonMultiplier + offsetKlmX, pentagon[2][1] * pentagonMultiplier + offsetKlmY),
          Person("Leszek Ludwicki", "Leszek", "KŁM", pentagon[3][0] * pentagonMultiplier + offsetKlmX, pentagon[3][1] * pentagonMultiplier + offsetKlmY),
          Person("Jachimek", "Szymon Jachimek", "KŁM", pentagon[1][0] * pentagonMultiplier + offsetKlmX, pentagon[1][1] * pentagonMultiplier + offsetKlmY),
          Person("Śliwiński", "Jakub Śliwiński", "KŁM", pentagon[4][0] * pentagonMultiplier + offsetKlmX, pentagon[4][1] * pentagonMultiplier + offsetKlmY),
          Person("Sherlock", "Sherlock", "SUM", pentagon[0][0] * pentagonMultiplier + offsetSumX, pentagon[0][1] * pentagonMultiplier + offsetSumY),
          Person("Oleksy", "Mariusz Oleksy", "SUM", pentagon[4][0] * pentagonMultiplier + offsetSumX, pentagon[4][1] * pentagonMultiplier + offsetSumY),
          Person("Kotowski", "Maciej Kotowski", "SUM", pentagon[2][0] * pentagonMultiplier + offsetSumX, pentagon[2][1] * pentagonMultiplier + offsetSumY),
          Person("Hirsz", "Żania", "SUM", pentagon[3][0] * pentagonMultiplier + offsetSumX, pentagon[3][1] * pentagonMultiplier + offsetSumY),
          Person("Kuba", "Kuba", "SUM", pentagon[1][0] * pentagonMultiplier + offsetSumX, pentagon[1][1] * pentagonMultiplier + offsetSumY),
          Person("Maciaszek", "Rock", "Popiół", hexagons[0][0] * multiplier + offsetPopX, hexagons[0][1] * multiplier + offsetPopY),
          Person("Seto", "Seto", "Popiół", hexagons[1][0] * multiplier + offsetPopX, hexagons[1][1] * multiplier + offsetPopY),
          Person("Górka", "Kromka", "Popiół", hexagons[4][0] * multiplier + offsetPopX, hexagons[4][1] * multiplier + offsetPopY),
          Person("Skorowski", "NieDaRadek", "Popiół", hexagons[3][0] * multiplier + offsetPopX, hexagons[3][1] * multiplier + offsetPopY),
          Person("Łosiu", "Łosiu", "Popiół", hexagons[2][0] * multiplier + offsetPopX, hexagons[2][1] * multiplier + offsetPopY),
          Person("Kamiński", "Jacek Kamiński", "Popiół", hexagons[5][0] * multiplier + offsetPopX, hexagons[5][1] * multiplier + offsetPopY),
          Person("Randall", "Randall", "GON", pentagon[2][0] * pentagonMultiplier + offsetGonX, pentagon[2][1] * pentagonMultiplier + offsetGonY),
          Person("Martin Stankiewicz", "Martin", "GON", pentagon[1][0] * pentagonMultiplier + offsetGonX, pentagon[1][1] * pentagonMultiplier + offsetGonY),
          Person("Kiełczykowski", "Wiktor Kiełczykowski", "GON", pentagon[0][0] * pentagonMultiplier + offsetGonX, pentagon[0][1] * pentagonMultiplier + offsetGonY),
          Person("Steifer", "Janek Steifer", "GON", pentagon[4][0] * pentagonMultiplier + offsetGonX, pentagon[4][1] * pentagonMultiplier + offsetGonY),
          Person("Ania Jaszczyk ", "Ania Jaszczyk", "GON", pentagon[3][0] * pentagonMultiplier + offsetGonX, pentagon[3][1] * pentagonMultiplier + offsetGonY),
          Person("Nemi", "Nemi", "Inne", -600, 10),
          Person("McGinn", "Bartosz (Walter McGinn)", "Inne", 570, 0),
          Person("Arek", "Arek Glinka", "Inne", 650, 75),
          Person("Stroch", "Mateusz Janik", "Inne", 570, 100),
          Person("Raslo", "Michał Bronk", "Inne", -380, -490),
          Person("Josh", "Josh Szpilarski", "Inne", -220, 500),
          Person("Ruciński", "Kacper Ruciński", "Inne", -120, 600),
          Person("Sławek (Emmet Dorman)", "Sławek (Emmet Dorman)", "Inne", 400, 400),
          Person("Sławek (Rupert Blakely)", "Sławek (Rupert Blakely)", "Inne", 500, 500),
          Person("Ewa (Cordelia Freeman)", "Ewa (Cordelia Freeman)", "Inne", 400, 600),
          Person("Iwona (Claudia Fillman)", "Iwona (Claudia Fillman)", "Inne", 300, 500),
          Person("Smok", "Patrycja Smok", "Inne", 470, -520),
          Person("Graba", "Graba", "Inne", 300, -520),
          Person("Szweda", "Jędrzej Szweda", "Inne", 530, -450),
          Person("Justyna Kapa (Olimpia Monroe)", "Justyna Kapa", "Inne", 550, -300),
          Person("Zuzanna Białogłowska (Zofia Mazurek)", "Zuzanna Białogłowska", "Inne", 650, -200),
          Person("Kasia Urbańska (Deborah Whittaker)", "Kasia Urbańska", "Inne", 550, -100),
          Person("Ewa Samulak (Elly Freeman)", "Ewa Samulak", "Inne", 450, -200)]

sessions = []

def setup():
    nameColumn = workSheet["D"]
    currentSession = Session()
    for nameCell in nameColumn:
        if (nameCell.value == None):
            if len(currentSession.people) > 0:
                currentSession.sessionLengthInSeconds += extractSessionLengthInSeconds(nameCell.row - 1)
                sessions.append(currentSession)
                currentSession = Session()
                currentSession.people = []
            continue
        found = False
        for person in people:
            if person.searchName in nameCell.value:
                found = True
                person.sessionCount += 1
                person.sessionLengthInSeconds += extractSessionLengthInSeconds(nameCell.row)
                sessionName = extractSessionName(nameCell.row)
                if currentSession.name is not sessionName and len(currentSession.people) > 0:
                    currentSession.sessionLengthInSeconds += extractSessionLengthInSeconds(nameCell.row - 1)
                    sessions.append(currentSession)
                    currentSession = Session()
                    currentSession.people = []
                currentSession.people.append(person)
                currentSession.name = sessionName
                currentSession.teamName = extractTeamName(nameCell.row)

        if found == False:
            print(nameCell.value)
            

def getPeopleHoursHtml():
    def sortPeople(person):
        return person.sessionLengthInSeconds

    people.sort(key=sortPeople)

    x = list(map(lambda x: x.displayName, people))
    y = list(map(lambda x: x.sessionLengthInHours(), people))
    hovertext = list(map(lambda x: x.sessionLengthString(), people))

    fig = go.Figure(data=[go.Bar(
        x=x,
        y=y,
        hovertext=hovertext,
        marker_color=_get_colors(len(x))
    )])

    fig.update_layout(title_text='BCU - Godziny', template="plotly_dark")
    fig.write_html('/Users/przemyslawszafulski/Developer/BaniakBaniaka/baniak_wszystko_godziny.html')

def getPeopleSessionsHtml():
    def sortPeople(person):
        return person.sessionCount

    people.sort(key=sortPeople)

    x = list(map(lambda x: x.displayName, people))
    y = list(map(lambda x: x.sessionCount, people))
    hovertext = list(map(lambda x: x.sessionCount, people))

    fig = go.Figure(data=[go.Bar(
        x=x,
        y=y,
        hovertext=hovertext,
        marker_color=_get_colors(len(x))
    )])

    fig.update_layout(title_text='BCU - Sesje', template="plotly_dark")
    fig.write_html('/Users/przemyslawszafulski/Developer/BaniakBaniaka/baniak_wszystko_sesje.html')

def getAverageSessionLength():
    totalTimeInSeconds = 0
    for session in sessions:
        totalTimeInSeconds += session.sessionLengthInSeconds
    averageTime = totalTimeInSeconds / len(sessions)
    print("Total session time is: " + str(totalTimeInSeconds) + "s. Average time is: " + str(averageTime) + "s.")

def getSessionLengthsHTML():
    hobo = []
    sum = []
    klm = []
    gon = []
    popiol = []

    hoboAverage = 0
    sumAverage = 0
    klmAverage = 0
    gonAverage = 0
    popiolAverage = 0

    def sessionLengthString(totalSeconds):
        total = totalSeconds
        seconds = total % 60
        total -= seconds
        minutes = (total / 60) % 60
        total -= minutes * 60
        hours = (total / 3600)
        string = ""
        if (hours < 10):
            string += "0" + str(int(hours))
        else:
            string += str(int(hours))
        string += ":"
        if (minutes < 10):
            string += "0" + str(int(minutes))
        else:
            string += str(int(minutes))
        string += ":"
        if (seconds < 10):
            string += "0" + str(int(seconds))
        else:
            string += str(int(seconds))
        return string

    
    for session in sessions:
        groupName = session.teamName
        if groupName == "HOBO":
            hobo.append(session)
            hoboAverage += session.sessionLengthInHours()
        elif groupName == "SUM":
            sum.append(session)
            sumAverage += session.sessionLengthInHours()
        elif groupName == "KŁM":
            klm.append(session)
            klmAverage += session.sessionLengthInHours()
        elif groupName == "GON":
            gon.append(session)
            gonAverage += session.sessionLengthInHours()
        elif groupName == "Popiół":
            popiol.append(session)
            popiolAverage += session.sessionLengthInHours()
    
    hoboAverage = hoboAverage / len(hobo)
    sumAverage = sumAverage / len(sum)
    gonAverage = gonAverage / len(gon)
    klmAverage = klmAverage / len(klm)
    popiolAverage = popiolAverage / len(popiol)
    hoboAverageHovertext = sessionLengthString(hoboAverage * 3600)
    sumAverageHovertext = sessionLengthString(sumAverage * 3600)
    gonAverageHovertext = sessionLengthString(gonAverage * 3600)
    klmAverageHovertext = sessionLengthString(klmAverage * 3600)
    popiolAverageHovertext = sessionLengthString(popiolAverage * 3600)

    fig = make_subplots()

    startPoint = 1
    const = 2
    x = list(range(startPoint, len(hobo) + startPoint))
    y = list(map(lambda x: x.sessionLengthInHours(), hobo))
    hovertext = list(map(lambda x: x.sessionLengthString(), hobo))
    fig.add_trace(
        go.Scatter(x=x, y=y, name="HOBO", hovertext=hovertext, line_color=hobo[0].color(), mode = 'lines')
    )
    
    fig.add_trace(
        go.Scatter(x=x, y=[hoboAverage] * len(x), hovertext = [hoboAverageHovertext] * len(x), name="HOBO - średnia", line_color=hobo[0].color(), line_dash="dash", mode = 'lines')
    )
    

    startPoint += len(x) + const
    x = list(range(startPoint, len(sum) + startPoint))
    y = list(map(lambda x: x.sessionLengthInHours(), sum))
    hovertext = list(map(lambda x: x.sessionLengthString(), sum))
    fig.add_trace(
        go.Scatter(x=x, y=y, name="SUM", hovertext=hovertext, line_color=sum[0].color(), mode = 'lines')
    )

    fig.add_trace(
        go.Scatter(x=x, y=[sumAverage] * len(x), hovertext = [sumAverageHovertext] * len(x), name="SUM - średnia", line_color=sum[0].color(), line_dash="dash", mode = 'lines')
    )

    startPoint += len(x) + const
    x = list(range(startPoint, len(klm) + startPoint))
    y = list(map(lambda x: x.sessionLengthInHours(), klm))
    hovertext = list(map(lambda x: x.sessionLengthString(), klm))
    fig.add_trace(
        go.Scatter(x=x, y=y, name="KŁM", hovertext=hovertext, line_color=klm[0].color(), mode = 'lines')
    )

    fig.add_trace(
        go.Scatter(x=x, y=[klmAverage] * len(x), hovertext = [klmAverageHovertext] * len(x), name="KŁM - średnia", line_color=klm[0].color(), line_dash="dash", mode = 'lines')
    )

    startPoint += len(x) + const
    x = list(range(startPoint, len(gon) + startPoint))
    y = list(map(lambda x: x.sessionLengthInHours(), gon))
    hovertext = list(map(lambda x: x.sessionLengthString(), gon))
    fig.add_trace(
        go.Scatter(x=x, y=y, name="GON", hovertext=hovertext, line_color=gon[0].color(), mode = 'lines')
    )

    fig.add_trace(
        go.Scatter(x=x, y=[gonAverage] * len(x), hovertext = [gonAverageHovertext] * len(x), name="GON - średnia", line_color=gon[0].color(), line_dash="dash", mode = 'lines')
    )

    startPoint += len(x) + const
    x = list(range(startPoint, len(popiol) + startPoint))
    y = list(map(lambda x: x.sessionLengthInHours(), popiol))
    hovertext = list(map(lambda x: x.sessionLengthString(), popiol))
    fig.add_trace(
        go.Scatter(x=x, y=y, name="Popiół", hovertext=hovertext, line_color=popiol[0].color(), mode = 'lines')
    )

    fig.add_trace(
        go.Scatter(x=x, y=[popiolAverage] * len(x), hovertext = [popiolAverageHovertext] * len(x), name="Popiół - średnia", line_color=popiol[0].color(), line_dash="dash", mode = 'lines')
    )

    
    fig.update_layout(title_text='BCU - Długość sesji', template="plotly_dark")
    fig.update_xaxes(visible=False, showticklabels=False)
    fig.update_yaxes(title="Godziny")
    fig.update_layout(yaxis_range=[0,6.5])
    fig.write_html('/Users/przemyslawszafulski/Developer/BaniakBaniaka/bcu_druzyny.html')

setup()
getPeopleHoursHtml()
getPeopleSessionsHtml()
getAverageSessionLength()
getSessionLengthsHTML()
baniak_network.createNetwork(sessions, people)