import pandas as pd
import numpy as np
import json
import os

sheet_names = ["S1", "S2", "S3", "S4", "S5", "S6"]

dfs = {}

for sheet in sheet_names:
    df = pd.read_excel(
        "Timetable Workbook - SUTT Task 1.xlsx", sheet_name=sheet, header=None
    )
    dfs[sheet] = df
    df.columns = [
        "COM COD",
        "COURSE NO.",
        "COURSE TITLE",
        "Lecture",
        "Practical",
        "Units",
        "SEC",
        "INSTRUCTOR-IN-CHARGE / Instructor",
        "ROOM",
        "DAYS & HOURS",
        "MIDSEM DATE & SESSION",
        "COMPRE DATE & SESSION",
    ]
    df.drop([0, 2], axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.drop(0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.fillna("", inplace=True)

    def get_section_type(sec_value):
        if sec_value.startswith("P"):
            return "Practical"
        elif sec_value.startswith("L"):
            return "Lecture"
        elif sec_value.startswith("T"):
            return "Tutorial"

    def split_str(s):
        a = s.split(" ")
        return a

    def time_slot(s):
        if s == '1':
            correct_slot = '[8,9]'
        elif s == '2':
            correct_slot = '[9,10]'
        elif s == '3':
            correct_slot = '[10,11]'
        elif s == '4':
            correct_slot = '[11,12]'
        elif s == '5':
            correct_slot = '[1,2]'
        elif s == '6':
            correct_slot = '[2,3]'
        elif s == '7':
            correct_slot = '[3,4]'
        elif s == '8':
            correct_slot = '[4,5]'
        elif s == '9':
            correct_slot = '[5,6]'
        return correct_slot


    def timeDict(s1, s2):

        timing_dict = {"day": None, "slots": None}
        timing = []

        s2_split = s2.split()
        
        slot_list = []

        if len(s2_split) > 1:
            i = 0
            while i < len(s2_split):
                slot_list.append(time_slot(s2_split[i]))
                i += 1
        else:
            slot_list.append(time_slot(s2_split[0]))
        

        index = 0
        while index < len(s1):
            timing.append(timing_dict.copy())
            index += 1

        index = 0
        while index < len(s1):
            timing[index]["day"] = s1[index]
            index += 1

        index = 0
        while index < len(s1):
            timing[index]["slots"] = slot_list
            index += 1

        return timing

    def sample_dict(index_start, index_last):
        for i in range(index_start, index_last):
            j = i
            instructor = []
            while j < index_last:
                instructor.append(df.loc[j, "INSTRUCTOR-IN-CHARGE / Instructor"])
                j += 1
            s = df.loc[i, "DAYS & HOURS"]

            timing = []
            s_split = s.split("  ")
            index = 0

            while index < len(s_split):
                timing = timing + timeDict(
                    split_str(s_split[index]), s_split[(index + 1)]
                )
                index += 2

            section_dict = {
                "section_type": get_section_type(df.loc[i, "SEC"]),
                "section_number": df.loc[i, "SEC"],
                "instructors": instructor,
                "room": str(df.loc[i, "ROOM"]),
                "timings": timing,
            }
            i += index_last

            return section_dict

    list = [0]

    i = 0

    while i < len(df) - 1:

        if df.loc[i + 1, "SEC"] != "":

            list.append(i + 1)

        i += 1
    list.append(len(df))

    section_dict_list = []
    i = 0
    while i < len(df):
        if (i + 1) == len(list):
            break
        section_dict_list.append(sample_dict(list[i], list[i + 1]))
        i += 1

    dict = {
        "course_code": df.loc[0, "COURSE NO."],
        "course_title": df.loc[0, "COURSE TITLE"],
        "credits": {
            "lecture": df.loc[0, "Lecture"],
            "practical": df.loc[0, "Practical"],
            "units": df.loc[0, "Units"],
        },
        "section": section_dict_list,
    }

    # with open('Simarjot.json', 'r') as json_file:
    #         data = json.load(json_file)
   
    # # Append the new data for the current sheet
    # data.append(dict)

    # # Write the updated data back to the JSON file

    
    json_output = json.dumps(dict, indent=4)
    # print(json_output)

    # with open('Simarjot.json', 'a') as json_file:
    #     json.dump(json_file, indent=4)

    # Open the file in append mode
    with open('output.json', 'a') as f:
        f.write(json_output + "\n")  # Add a newline for readability between entries

    
    
