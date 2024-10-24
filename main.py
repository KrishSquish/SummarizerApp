import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk

root = tk.Tk()
root.geometry('500x500')
root.resizable(False, False)

def select_file():
    global data
    filetypes = (
        ('Excel files', '*.xlsx'),
    )

    file = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    file = file.replace("/","\\")
    fileText.config(text=file)
    data = (pd.read_excel(file)).values.tolist()


def getData():
    global data
    ampScore = 0
    speakerScore = 0
    teamData = []

    team = int(submitTeamNum.get())
    for i in range(len(data)):
        if data[i][2] == team:
            print("Team in", i, data[i][2])
            teamData.append(data[i])

    #printing stats
    for i in teamData:
        #avg amp scores
        ampScore += i[6]
        speakerScore += i[7]

    ampScore /= len(teamData)
    speakerScore /= len(teamData)

    print("Average Amp Scored:",ampScore)
    print("Average Speaker Scored:",speakerScore)



title = tk.Label(
    text= "Summarizer App",
    fg ="black"
)

# open button
openFile = ttk.Button(
    root,
    text='Open Spreadsheet File',
    command=select_file
)

fileText = tk.Label(
    text = "No File Selected"
)

enterTeams = tk.Label(
    text= "Which team would you like data for?",
    fg ="black"
)
submitTeamNum = tk.Entry()
button = tk.Button(text = "Get Data", command = getData)

title.pack()
fileText.pack()
openFile.pack()
enterTeams.pack()
submitTeamNum.pack()
button.pack()

root.mainloop()
