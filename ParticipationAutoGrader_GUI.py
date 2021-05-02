#!/usr/bin/env python
'''
Participation Auto Grader for INST354
'''
import sys
import tkinter
from tkinter import ttk
import pandas as pd
from difflib import SequenceMatcher
import statistics as st
import json
import time
import os
if sys.version_info[0] >= 3:
    import PySimpleGUI as sg
else:
    import PySimpleGUI27 as sg

print = sg.EasyPrint

studentDictionary = {}
studentArr = []
studentNameArr = []
studentsArr = []
selfGradeArr = []
selfRawArr = []
peerGradeArr = []
peerRawArr = []
sg.theme('DefaultNoMoreNagging')


def mainWin():
    form_rows = [
        [sg.Text("Fill in the required information and press start")],
        [sg.Text('Select Class Roster (.txt)', size=(25, 1)),
         sg.InputText(key='classRoster'), sg.FileBrowse()],
        [sg.Text('Select Participation File (.xlsx)', size=(25, 1)),
         sg.InputText(key='partFile'), sg.FileBrowse()],
        [sg.Button('Start'), sg.Exit()],
        [sg.Text('                                                                                                      Python 3.x - Jonathan Obenland', text_color='green')]]

    window = sg.Window("Participation Auto Grader")
    event, values = window.Layout(form_rows).Read()

    # start the event listener
    while True:
        if event is None or event == 'Exit':
            break
        if event == 'Start':
            classRoster = values['classRoster']
            partFile = values['partFile']
            if classRoster == '' or partFile == '':
                window.Close()
                sg.PopupError("Insufficent data or Null Pointers given")
                mainWin()
            elif classRoster != '' and partFile != '':
                window.Close()
                print("Auto Grader v1.0.0 by Jonathan Obenland")
                print("Populating class roster...")
                numStudents = populateDict(classRoster)
                print("Populating JSON index with", numStudents,
                      "students. Gathering peer review data...")
                gatherReviews(partFile)
                print("Complete! Awaiting output information...")
                output = sg.popup_get_folder("Enter Path to Save")
                os.chdir(output)
                # try:
                with open("output.json", "w") as outfile:
                    json.dump(studentDictionary, outfile)
                toXlsx = pd.DataFrame(studentDictionary).transpose()
                toXlsx.to_excel(r'output.xlsx', index=True)
                sg.Popup('Complete')
                # except:
                #     sg.Popup('ERROR! Close file and try agian')
            break


def main():
    rosterPath = input("Enter txt file containing roster: ")
    partFile = input("Enter xlsx file contatining participation: ")
    populateDict(rosterPath)
    gatherReviews(partFile)
    with open("output.json", "w") as outfile:
        json.dump(studentDictionary, outfile)


def populateDict(rosterPath):
    numStudents = 0
    with open(rosterPath) as file_in:
        for name in file_in:
            numStudents += 1
            name = name.replace("\n", "")
            studentArr.append(name)
            studentDictionary[name] = {
                "Name": name,
                "SelfScore": -1,
                "PeerScore": -1,
                "TotalScore": -1,
                "Students": 0,
                "SelfRaw": [],
                "PeerRaw": [],
                "TotalRaw": [],
                "Deduction": -1
            }
    return numStudents


def gatherReviews(partFile):
    df = pd.read_excel(partFile)
    print("Imported", len(df.index), "submissions")
    for i in range(len(df.index)):
        firstMember = df.iloc[i, 1:10]
        secondMember = df.iloc[i, 10:19]
        thirdMember = df.iloc[i, 19:28]
        fourthMember = df.iloc[i, 28:37]
        fifthMember = df.iloc[i, 37:46]
        addToDict(firstMember, True)
        addToDict(secondMember, False)
        addToDict(thirdMember, False)
        addToDict(fourthMember, False)
        addToDict(fifthMember, False)


def addToDict(series, isSelf):
    try:
        scoreArr = []
        studentName = series[0] + " " + series[1]
        studentName = str(studentName)
        for index in range(len(studentArr)):
            similarity = similar(studentArr[index], studentName)
            if similarity > 0.75:
                for item in series[2:8]:
                    scoreArr.append(item)
                    studentDictionary[studentArr[index]
                                      ]["TotalRaw"].append(int(item))
                    if isSelf:
                        studentDictionary[studentArr[index]
                                          ]["SelfRaw"].append(int(item))
                        studentDictionary[studentArr[index]
                                          ]["Deduction"] = 0
                    else:
                        studentDictionary[studentArr[index]
                                          ]["PeerRaw"].append(int(item))
                if isSelf:
                    studentDictionary[studentArr[index]
                                      ]["SelfScore"] = float(st.mean(scoreArr))
                else:
                    studentDictionary[studentArr[index]
                                      ]["PeerScore"] = float(st.mean(studentDictionary[studentArr[index]
                                                                                       ]["PeerRaw"]))
                    studentDictionary[studentArr[index]
                                      ]["Students"] = studentDictionary[studentArr[index]
                                                                        ]["Students"]+1
                studentDictionary[studentArr[index]
                                  ]["TotalScore"] = float(st.mean(studentDictionary[studentArr[index]
                                                                                    ]["TotalRaw"])) + studentDictionary[studentArr[index]
                                                                                                                        ]["Deduction"]
                break
            elif similarity <= 0.75 and similarity > 0.57:
                cont = sg.PopupYesNo(
                    "Is", studentArr[index], "and", studentName, "the same person?")
                if cont == "Yes":
                    for item in series[2:8]:
                        scoreArr.append(item)
                        studentDictionary[studentArr[index]
                                          ]["TotalRaw"].append(int(item))
                        if isSelf:
                            studentDictionary[studentArr[index]
                                              ]["SelfRaw"].append(int(item))
                            studentDictionary[studentArr[index]
                                              ]["Deduction"] = 0
                        else:
                            studentDictionary[studentArr[index]
                                              ]["PeerRaw"].append(int(item))
                    if isSelf:
                        studentDictionary[studentArr[index]
                                          ]["SelfScore"] = float(st.mean(scoreArr))
                    else:
                        studentDictionary[studentArr[index]
                                          ]["PeerScore"] = float(st.mean(studentDictionary[studentArr[index]
                                                                                           ]["PeerRaw"]))
                        studentDictionary[studentArr[index]
                                          ]["Students"] = studentDictionary[studentArr[index]
                                                                            ]["Students"]+1
                    studentDictionary[studentArr[index]
                                      ]["TotalScore"] = float(st.mean(studentDictionary[studentArr[index]
                                                                                        ]["TotalRaw"])) + studentDictionary[studentArr[index]
                                                                                                                            ]["Deduction"]
                    break
    except:
        studentName = "ERROR"


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


mainWin()
