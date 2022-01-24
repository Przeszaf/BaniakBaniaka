from openpyxl import load_workbook
import plotly.graph_objects as go
import numpy as np
import colorsys

workBook = load_workbook('/Users/przemyslawszafulski/Developer/BaniakBaniaka/Baniakowe erpegi.xlsx')
workSheet = workBook["Arkusz1"]

timeColumn = workSheet["F"]

def extractSessionLengthInSeconds(value):
    return (value.hour * 60 + value.minute) * 60 + value.second

def sessionLengthString(total):
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

def sessionLengthInHours(sessionLengthInSeconds):
        return sessionLengthInSeconds / 3600

sezon4 = 0
sezon5 = 0
sezon6 = 0
sezon7 = 0
sezon8 = 0
sezon9 = 0
for column in timeColumn:
    if column.row > 1:
        if column.row <= 28:
            sezon4 += extractSessionLengthInSeconds(column.value)
        elif column.row <= 100:
            sezon5 += extractSessionLengthInSeconds(column.value)
        elif column.row <= 154:
            sezon6 += extractSessionLengthInSeconds(column.value)
        elif column.row <= 234:
            sezon7 += extractSessionLengthInSeconds(column.value)
        elif column.row <= 299:
            sezon8 += extractSessionLengthInSeconds(column.value)
        elif column.row <= 329:
            sezon9 += extractSessionLengthInSeconds(column.value)

s4 = sezon4 / 27
s5 = sezon5 / 72
s6 = sezon6 / 54
s7 = sezon7 / 80
s8 = sezon8 / 65
s9 = sezon9 / 30


x = [4, 5, 6, 7, 8, 9]
y = list(map(lambda x: sessionLengthInHours(x), [s4, s5, s6, s7, s8, s9]))
hovertext = list(map(lambda x: sessionLengthString(x), [s4, s5, s6, s7, s8, s9]))

fig = go.Figure(data=[go.Bar(
    x=x,
    y=y,
    hovertext=hovertext,
    marker_color=_get_colors(len(x))
)])

fig.update_yaxes(title="Czas")
fig.update_xaxes(title="Sezon")
fig.update_layout(title_text='Średni czas sesji', template="plotly_dark")
fig.write_html('/Users/przemyslawszafulski/Developer/BaniakBaniaka/baniak_sredni_czas_sesji.html')