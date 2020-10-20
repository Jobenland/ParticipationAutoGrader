import pandas as pd
from difflib import SequenceMatcher
import statistics as st
import json

studentDictionary = {}
studentArr = []


def main():
    rosterPath = input("Enter txt file containing roster: ")
    partFile = input("Enter xlsx file contatining participation: ")
    populateDict(rosterPath)
    gatherReviews(partFile)
    with open("output.json", "w") as outfile:
        json.dump(studentDictionary, outfile)


def populateDict(rosterPath):
    with open(rosterPath) as file_in:
        for name in file_in:
            name = name.replace("\n", "")
            studentArr.append(name)
            studentDictionary[name] = {
                "Name": name,
                "Score": -1,
                "Students": -1,
                "Raw": [],
            }


def gatherReviews(partFile):
    df = pd.read_excel(partFile)
    for i in range(46):
        firstMember = df.iloc[i, 1:10]
        secondMember = df.iloc[i, 10:19]
        thirdMember = df.iloc[i, 19:28]
        fourthMember = df.iloc[i, 28:37]
        fifthMember = df.iloc[i, 37:46]
        addToDict(firstMember)
        addToDict(secondMember)
        addToDict(thirdMember)
        addToDict(fourthMember)
        addToDict(fifthMember)


def addToDict(series):
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
                                      ]["Raw"].append(int(item))
                studentDictionary[studentArr[index]
                                  ]["Score"] = float(st.mean(scoreArr))
                studentDictionary[studentArr[index]
                                  ]["Students"] += 1
                break
    except:
        studentName = "ERROR"


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


main()
