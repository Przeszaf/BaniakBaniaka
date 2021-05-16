from mimetypes import init
from openpyxl import load_workbook
from openpyxl.descriptors.base import String
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import colorsys
import random

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
    def __init__(self, searchName, displayName):
        self.searchName = searchName
        self.displayName = displayName
    
    searchName = ""
    displayName = ""
    sessionCount = 0
    sessionLengthInSeconds = 0
    
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
        
        

workBook = load_workbook('/Users/przemyslaw.szafulski/Downloads/Sesje_Caly_Sezon.xlsx')
workSheet = workBook["BCU do uzupełnienia"]
nameColumn = workSheet["D"]

def extractSessionLengthInSeconds(row):
    timeColumn = workSheet["G"]
    i = 1
    while True:
        value = timeColumn[row - i].value
        if (value != None):
            return (value.hour * 60 + value.minute) * 60 + value.second
        else:
            i += 1
            
def extractSessionDate(row):
    dateColumn = workSheet["F"]
    i = 1
    while True:
        value = dateColumn[row - i].value
        if (value != None):
            print(value.day)
            return value
        else:
            i += 1
    

people = [Person("Nikos", "Nikos"), 
          Person("Dzieran", "Dzieran"), 
          Person("Karol", "Karol"), 
          Person("Lewy", "Lewy"), 
          Person("Kozi", "Kozi"), 
          Person("Słowik", "Słowik"), 
          Person("Tremiszewski", "Wojciech Tremiszewski"), 
          Person("Giza", "Abelard Giza"),
          Person("Leszek", "Leszek"),
          Person("Jachimek", "Szymon Jachimek"),
          Person("Śliwiński", "Jakub Śliwiński"),
          Person("Maciaszek", "Rock"),
          Person("Oleksy", "Maciek Oleksy"),
          Person("Kotowski", "Maciej Kotowski"),
          Person("Hirsz", "Żania"),
          Person("Dairy", "Kuba"),
          Person("Seto", "Seto"),
          Person("Kromka", "Kromka"),
          Person("Skorowski", "NieDaRadek"),
          Person("Stroch", "Mateusz Janik"),
          Person("Glinka", "Arek Glinka"),
          Person("Łosiu", "Łosiu"),
          Person("Randall", "Randall"),
          Person("Martin", "Martin"),
          Person("Sherlock", "Sherlock"),
          Person("Nemi", "Nemi"),
          Person("Kamiński", "Jacek Kamiński"),
          Person("McGinn", "Bartosz (Walter McGinn)"),
          Person("Raslo", "Michał Bronk"),
          Person("Kiełczykowski", "Wiktor Kiełczykowski"),
          Person("Smok", "Patrycja Smok"),
          Person("Josh", "Josh Szpilarski"),
          Person("Steifer", "Janek Steifer"),
          Person("Sławek (Emmet Dorman)", "Sławek (Emmet Dorman)"),
          Person("Sławek (Rupert Blakely)", "Sławek (Rupert Blakely)"),
          Person("Szweda", "Jędrzej Szweda"),
          Person("Ewa (Cordelia Freeman)", "Ewa (Cordelia Freeman)"),
          Person("Iwona (Claudia Fillman)", "Iwona (Claudia Fillman)"),
          Person("Anna Jażdżyk ", "Anna Jażdżyk"),
          Person("Ruciński", "Kacper Ruciński")]


for nameCell in nameColumn:
    if (nameCell.value == None):
        continue
    found = False
    for person in people:
        if person.searchName in nameCell.value:
            found = True
            person.sessionCount += 1
            person.sessionLengthInSeconds += extractSessionLengthInSeconds(nameCell.row)
    if found == False:
        print(nameCell.value)
            
            
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

fig.update_layout(title_text='Całe BCU - Ilość Godzin', template="plotly_dark")
fig.write_html('/Users/przemyslaw.szafulski/Downloads/baniak_wszystko_godziny.html')