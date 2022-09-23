# Project Name: Comp Grid Automation
# Author: Jiajian(Drew) Chen
# Coaches: Edward, Stella
# Last Edit: 08192022


# import packages needed
import pandas as pd
import numpy as np
import warnings
import sidetable
import os
from pywebio.input import *
from pywebio.output import *
from pywebio import start_server
warnings.filterwarnings("ignore")

# **************** Directory Path below is the only part for modification by users ***************
# * Working Directory
# for Windows users
os.chdir("C:\\Users\\mdjij\\OneDrive - IRI\\Desktop\\Automation\\Final Presentation\\Molson Coors")

# for MacOS users
# os.chdir("/Users/chenjiajian/IRI - Initiative/Comp Grid")

# ************************************************************************************************

def main():

    # Title and instruction links
    put_markdown("# **Comp Grid Automation**")
    put_markdown("""
    Create Comp Grid by just a few clicks!!!  
    Please make sure you already read [README](https://iriworldwide.sharepoint.com/:t:/r/sites/grpstrategicanalytics/Shared%20Documents/Analytics%20Modeling%20and%20Engagement%20(AME)%20Team/Code%20Tools/Comp%20Grid%20Automation%20Python/ReadMe.txt?csf=1&web=1&e=eTLu4y) and follow the instructions on [Documentation](https://iriworldwide.sharepoint.com/:p:/r/sites/grpstrategicanalytics/Shared%20Documents/Analytics%20Modeling%20and%20Engagement%20(AME)%20Team/Code%20Tools/Comp%20Grid%20Automation%20Python/Documentation%20of%20Comp%20Grid%20Automation.pptx?d=w75c205661e1642c8b0caa68c9439411e&csf=1&web=1&e=jv1VsX).
    [Templates for input file can be found here](https://iriworldwide.sharepoint.com/:f:/r/sites/grpstrategicanalytics/Shared%20Documents/Analytics%20Modeling%20and%20Engagement%20(AME)%20Team/Code%20Tools/Comp%20Grid%20Automation%20Python/Templates?csf=1&web=1&e=wq7KGg).
    You can refresh (F5) this page, without restarting the code, if you want to redo.
    """)

    # load the input file
    INPUT = input_group("Upload", [
        file_upload("PPGs & Attributes.xlsx", name="InputPath", required=True),
        file_upload("Interaction Rules.xlsx", name="LogicPath", required=True),
        file_upload("AMP PPG list (*OPTIONAL*)", name = "PPGList", required= False, help_text= "The PPG list downloaded directly from AMP")
    ])

    # convert the file we read to dataframe
    df = pd.read_excel(INPUT["InputPath"]["filename"])
    logic = pd.read_excel(INPUT["LogicPath"]["filename"], sheet_name=None)

    # extract ppg ids and ppg names of AMP PPG List, if uploaded
    if INPUT['PPGList'] != None:
        amp = pd.read_csv(INPUT["PPGList"]["filename"])
        amp = amp[["PPG Id", "PPG Name"]]
        amp = amp.set_index("PPG Name")
        amp.drop_duplicates(inplace=True)

    # store key parameters of input file
    PPG_col = df.columns[0] # the name of the ppg column
    PPGs = df[PPG_col].unique() # a set of ppg without duplicates

    # sorting
    IsSort = actions("Alphabetically sort the target PPGs?", ["Yes","No"])
    if IsSort == "Yes":
        df_target = df[df["IsTarget"] == 1].sort_values("PPG Name")
        df_comp = df[df["IsTarget"] == 0]
        df_sorted = pd.concat([df_target,df_comp], ignore_index= True)
    else:
        df_sorted = df

    # if the input is correctly formatted, proceed
    if set(df.columns[2:]) == set(logic.keys()):
        for k in logic.keys():
            logic[k] = logic[k].set_index(f"{k}") # where we need A1 cell of rule matrix aligned with tab name
        df_merged = merger(df_sorted)
        method = select(label="Which method to use?", options=["IRE", "Score"])
        list_process = 0
        if method == "Score":
            while list_process == 0:
                popup("Generating Comp Grid Output...")
                df_output, list_process = list_output_score(df_merged, PPG_col, logic)
            close_popup()
        elif method == "IRE":
            while list_process == 0:
                popup("Processing Comp Matrix Output...")
                df_output, list_process = list_output_IRE(df_merged, PPG_col, logic)

        # formatting and summary table
        if method == "Score":
            df_output =  format_and_statistics(df_output, PPG_col, "Score")
        elif method == "IRE":
            df_output =  format_and_statistics(df_output,PPG_col,"comptype")


        # PPG Id population
        if INPUT['PPGList'] != None:
            for r in range(len(df_output)):
                df_output.loc[r, "PPGID"] = amp.loc[df_output.loc[r,"PPG"],"PPG Id"]
                df_output.loc[r, "CompPPGID"] = amp.loc[df_output.loc[r, "CompPPG"], "PPG Id"]

        df_output = df_output[df_output.iloc[:,-1] != "X"]
        if method == "IRE":
            df_output.iloc[:,-1] = df_output.iloc[:,-1].replace(["2","1","0"],["Expected", "Regular", "Ignore"])
        df_output.to_excel("CompGrid_AMP.xlsx", index = False) # for scope defining
        df_output.to_csv("CompGrid_AMP_csv.csv", index = False) # output csv
        close_popup()
        put_html('<hr>')
        put_markdown("\n **Now you may find list_output.xlsx in your folder**")

        grid_process = 0
        while grid_process == 0:
            popup("Generating Matrix Output...")
            grid, grid_process = grid_output(df_merged, PPG_col, PPGs)
        grid.to_excel("CompMatrix.xlsx") # output - clean grid
        close_popup()
        put_markdown("\n **Now you may find matrix_output.xlsx in your folder** ")

        put_markdown("*The End*")
    else:
        put_markdown("The Input Attribute Names are not aligned with Logic Tab Names",)

def list_output_score(df, PPG, logic):
    df['Interaction'] = 0
    for r in range(len(df)):
        if df.loc[r, PPG + "_x"] == df.loc[r, PPG + "_y"]:
            df.loc[r, "Interaction"] = "X"
        else:
            for k in list(logic.keys()):
                df.loc[r,"Interaction"] += logic[k].loc[df.loc[r,k+"_x"], df.loc[r,k+"_y"]]
    return df, 1

def list_output_IRE(df, PPG, logic):
    df['Interaction'] = np.nan
    for r in range(len(df)):
        if df.loc[r, PPG + "_x"] == df.loc[r, PPG + "_y"]:
            df.loc[r, "Interaction"] = "X"
        else:
            condition_expected = True
            condition_ignore = list()
            for k in list(logic.keys()):
                condition_expected = condition_expected & (logic[k].loc[df.loc[r,k+"_x"], df.loc[r,k+"_y"]] == "Strong")
            for k in list(logic.keys()):
                condition_ignore.append(pd.isnull(logic[k].loc[df.loc[r,k+"_x"], df.loc[r,k+"_y"]]))
            if condition_expected == True:
                df.loc[r,"Interaction"] = "2"
            elif sum(condition_ignore) > 0:
                df.loc[r,"Interaction"] = "0"
            else:
                df.loc[r,"Interaction"] = "1"
    return df, 1

def format_and_statistics(df, PPG_col, interaction_col):
    #formatting
    df['PPGID'] = np.nan
    df['CompPPGName'] = np.nan
    df = df[[PPG_col + '_x', 'PPGID', PPG_col + '_y', "CompPPGName", "Interaction"]]
    # change Interaction to specific name
    df.columns = ["PPG", "PPGID", "CompPPG", "CompPPGID", interaction_col]  # readability

    Summary_Table = df.stb.freq([interaction_col])
    Summary_Table['percent'] = [f"{round(i, 2)} %" for i in Summary_Table['percent']]
    Summary_Table['cumulative_percent'] = [f"{round(i, 2)} %" for i in Summary_Table['cumulative_percent']]
    put_html(Summary_Table.to_html(border=0, index=False)).send()

    if interaction_col == "Score":
        df_noX = df[df[interaction_col] != "X"][interaction_col]
        description = df_noX.astype(int).describe( percentiles= [.33,.66])
        description = pd.DataFrame(round(description, 2))
        put_html(description.to_html(border=0)).send()
    return df

def merger(df):
    df1 = df[df['IsTarget']==1].copy()
    df0 = df.copy()
    df1['key'] = 1
    df0['key'] = 1
    # merge for all combos
    df_merged = df1.merge(df0, on = "key").drop(['key','IsTarget_x','IsTarget_y'], axis = 1)
    return df_merged

def grid_output(df, PPG, PPGs):
    grid = pd.DataFrame(np.zeros((len(df[PPG+"_x"].unique()),len(df[PPG+"_y"].unique()))), index = df[PPG+"_x"].unique(), columns = df[PPG+"_y"].unique())
    for i in grid.index:
        grid.loc[i] = list(df[df[PPG+"_x"] == i]["Interaction"])
    return grid, 1

def Score_to_IRE(d1, d2):
    rule01 = df['Score']<d1
    rule12 = df["Score"]>d2
    df.loc[rule01, "Score"] = 0
    df.loc[rule02, "Score"]


if __name__ == "__main__":
    start_server(main, port=8080, debug=True, auto_open_webbrowser= True)
