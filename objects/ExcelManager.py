from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pyautogui
from pathlib import Path
import time
import pandas as pd
from selenium.webdriver.common.keys import Keys

class ExcelManager:
    SCORES_LIST = [750,500,250,100,100,100,100,100,100,100,100,100]
    #define the global variable "Scores List"
    def __init__(self,df,leaderboard_df,link=None,booster=1):
        self.df = df
        self.leaderboard_df = leaderboard_df
        self.booster = booster
        # self.answer = answer
        # self.link = link


    @classmethod
    def read_excel_file(cls,path,path2):
        """
        read_excel_file(cls,path)

        Description: 
        Prints out the values from every cell within the rowstart,rowend,colstart,colend parameters, from a workbook that can be found via path. This also serves as a constructor for the class ExcelManager

        Parameters:
            Path(str):
            path to the workbook
            
            rowstart,rowend(int):
            indicate what rows to read from
            
            colstart,colend(int): 
            indicate what columns to read from

        """
        # wb = load_workbook(path)
        # ws = wb.active
        # for row in range(rowstart,rowend):
        #     for column in range(colstart,colend):
        #     col = get_column_letter(column)
        #     print(ws["{}{}".format(col,row)].value)
        #     print("=====================================")
        df = pd.read_excel(path)
        leaderboard_df = pd.read_excel(path2)
        # answer = int(input("Please input the answer to this weeks problem"))
        # link = str(input("Please link the page to the current leaderboard"))

        return cls(df,leaderboard_df,booster=1)

    @staticmethod
    def clean_names(origional_name):
        """
        clean_names(origional_name)

        Description:
        Removes any unwanted whitespace from a name
        I.e: " Juri Mikoshiba " turns into "Juri Mikoshiba"

        Parameters:
            origional_name(str):
            string that you want to parse 


        """
        split_on_whitespace = origional_name.split(" ")
        print(split_on_whitespace)
        for i in range(len(split_on_whitespace)):
            x = split_on_whitespace[i]
            x = x.replace("\t","")
            x = x.replace("\n","")
            x = x.replace("(♦)","")
            x = x.replace("[♦♦]","")
            split_on_whitespace[i] = x
            
        
        print(i)

        name = [i for i in split_on_whitespace if i != ""]
        cleaned_name = " ".join(name)

        return cleaned_name

    @staticmethod
    def find_newly_downloaded_files():
        current_path = Path(".")
        weekly_scores = None
        weekly_leaderboard = None
        answer_detection = current_path.resolve() / "AnswerDetection"
        excel_files = list(answer_detection.glob("*.xlsx"))
        print(excel_files)
        for file in excel_files:
            print(file.as_posix())
            if("Math Commitee Week" in file.as_posix()):
                weekly_scores = file.as_posix()
            elif ("Leaderboard" in file.as_posix()):
                #file containing the data for the current leaderboard
                weekly_leaderboard = file.as_posix()
            if weekly_leaderboard != None and weekly_scores != None:   
                return (weekly_scores,weekly_leaderboard)
            else:
                raise OSError("No Leaderboard/New weekly data found in file {}".format(str(answer_detection)))

    def clean_df(self):
        df = self.df
        df["Name"] = df.apply(lambda x: ExcelManager.clean_names(x["Name"]),axis=1)
        df["Score"] = df.apply(lambda x: ExcelManager.clean_scores(x["Score"]),axis=1)
        df["Class"] = df.apply(lambda x: ExcelManager.clean_grade_string(x["Class"]),axis=1)

        return df

    @staticmethod
    def clean_scores(origional_score):
        #removes any unwanted whitespace from a score
        if type(origional_score) == str:
            #if the score is a string
            #process a string (I.e: turn a score from " 750 " into 750)
            origional_score = int(origional_score)

        return origional_score

    @staticmethod
    def clean_grade_string(string):
        #code to clean the strings for grade values i.e: " 8A " becomes "G8A"

        #code to replace extra whitespaces
        string = string.replace(" ","")
        if "G" not in string:
            string = "G" + string
        return string

    

    def process_results_file(self):
        df1 = self.df
        df1 = df1.dropna(axis=1)
        df1 = df1.drop(["Email","Start time","ID","(Optional) Your Working in text - It may be featured in the Daily Notices as the solution of the week!"],axis=1)
        df1 = df1.drop_duplicates(subset="Name (Please type your name)",keep="last")
        return df1

    def check_answers_and_assign_scores(self):
        df = self.df
        correct_df_1 = df[df["Your answer"] == self.answer]

        #rename the "Name" column of the dataframe containing the info about the people who got the correct answers
        correct_df = correct_df_1.rename({"Name (Please type your name)":"Name","Homeroom":"Class"},axis=1)

        #main code to assign the scores
        correct_df["Completion time"] = pd.to_datetime(correct_df["Completion time"])
        correct_df = correct_df.sort_values("Completion time")
        scores_to_be_used = ExcelManager.SCORES_LIST[0:len(correct_df)] * self.booster
        correct_df["Score"] = scores_to_be_used
        #as we no longer need the completion time and the answer columns anymore, drop them
        correct_df = correct_df.drop(["Completion time","Your answer"],axis=1)
        correct_df = correct_df.reset_index()
        correct_df = correct_df.drop("index",axis=1)
        print(correct_df)
        return correct_df

    def check_and_remove_descrepancies(self):
        df = self.leaderboard_df
        df2 = self.df
        indexes = list(df["Name"])
        indexes_new = list(df2["Name"])
        df = df.set_index("Name")
        df2 = df2.set_index("Name")
        print(indexes)
        print(indexes_new)
        print(df)
        print(df2)
        for i in range(len(indexes_new)):
            #loop through the newly answered questions
            if indexes_new[i] in indexes:
                #the same person has answered another question
                df.loc[indexes_new[i],"Score"] = df.loc[indexes_new[i],"Score"] + df2.loc[indexes_new[i],"Score"]

            else:
            
                #a new person has answered
                print(df2.loc[indexes_new[i]])
                df = df.append(df2.loc[indexes_new[i]],ignore_index=False)
        df = df.sort_values("Score",ascending=False)
        return df

    def update_leaderboard(self,answer,booster=1):
        self.leaderboard_df = self.clean_df(self.leaderboard_df)
        self.leaderboard_df = self.clean_names(self.leaderboard_df)
        # df1 = clean_df(df1)
        
        #after we have processed the leaderboard dataframe, process the dataframe containing the new submissions
        self.df = self.process_results_file(self.df)
        self.df = self.check_answers_and_assign_scores(self.df,answer=answer,booster=booster)
        self.df = ExcelManager.clean_df(self.df)
        print(self.df)

        final_df = self.check_and_remove_descrepancies(self.leaderboard_df,self.df)
        #it is now done!!1
        return final_df

    def check_answers_and_assign_scores(self):
        # df,answer,booster=1
        df = self.df
        correct_df_1 = df[df["Your answer"] == self.answer]

        #rename the "Name" column of the dataframe containing the info about the people who got the correct answers
        correct_df = correct_df_1.rename({"Name (Please type your name)":"Name","Homeroom":"Class"},axis=1)

        #main code to assign the scores
        correct_df["Completion time"] = pd.to_datetime(correct_df["Completion time"])
        correct_df = correct_df.sort_values("Completion time")
        scores_to_be_used = ExcelManager.SCORES_LIST[0:len(correct_df)] * self.booster
        correct_df["Score"] = scores_to_be_used
        #as we no longer need the completion time and the answer columns anymore, drop them
        correct_df = correct_df.drop(["Completion time","Your answer"],axis=1)
        correct_df = correct_df.reset_index()
        correct_df = correct_df.drop("index",axis=1)
        print(correct_df)
        return correct_df


    def modify_score(self,path,value,name,cellstr=None,df=None):
        """
        Docstring:
        modify one value in the excel spreadsheet and save it

        Parameters:
            Path:
            Path to the excel file that you want to modify
            
            cellstr:
            Index of the cell
            
            Value: 
            Value that you want to replace with
            
            name:
            Name of person that has thier score being modified


        """
        wb = self.load_workbook(path)
        if cellstr != None:
            wb[cellstr].value = value
            print("Modified Value of Cell {} to: {}".format(cellstr,value))
            wb.save("Math Commitee Data modified.xlsx")
        else:
            #we do not know the precise location of the cell to be edited, so thus we will edit the dataframe and save that instead.
            if df == None:
            #no df provided, trigger error
                raise ValueError("Please provide a valid dataframe, or provide a cellstring")

            else:
                #edit the dataframe
                df.loc[name,"Score"] = value
                print("{}'s score has been changed to {}!".format(name,value))
                return df

