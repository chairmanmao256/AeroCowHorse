import numpy as np
import pandas as pd
import os

def preprocessing(fname:str)->pd.DataFrame:
    df = pd.read_excel(fname)
    df = df[df["基本信息"].notna()]
    name_df = df.loc[df["基本信息"]=="姓名"]
    day_dict = {
        "每日统计":     name_df["每日统计"].to_numpy()[0][0:2],
        "Unnamed: 3": name_df["Unnamed: 3"].to_numpy()[0][0:2],
        "Unnamed: 4": name_df["Unnamed: 4"].to_numpy()[0][0:2],
        "Unnamed: 5": name_df["Unnamed: 5"].to_numpy()[0][0:2],
        "Unnamed: 6": name_df["Unnamed: 6"].to_numpy()[0][0:2],
        "Unnamed: 7": name_df["Unnamed: 7"].to_numpy()[0][0:2],
        "Unnamed: 8": name_df["Unnamed: 8"].to_numpy()[0][0:2],
    }
    df.rename(columns=day_dict, inplace = True)
    df.fillna("NO INFO", inplace=True)
    df.drop(0, inplace = True)
    
    return df

def check_late(df: pd.DataFrame, hour = 9, min = 5,
               dayList = ["周一","周二","周三","周四", "周五"]) -> int:
    late_num = 0
    for day in dayList:
        time_str = df[day].to_numpy()[0]
        if time_str!="NO INFO":
            hour_, min_ = time_str.split(":")
            if (int(hour_) == hour and int(min_) > min) or (int(hour_) > hour):
                late_num += 1
    
    return late_num

def calc_bonus(time: np.ndarray):
    return np.maximum(np.minimum(10.0, (np.round(time, 0)-45)), 0.0) * 10.0 + np.minimum((np.round(time, 0) - 45),0.0) * 20.0

if __name__=="__main__":
    files = [f for f in os.listdir() if os.path.isfile(f)]


    # sort the file names according to the date
    # After sorting, the increasing order in index means increasing order in date
    tables = []
    dates = []
    for f in files:
        name, ext = os.path.splitext(f)
        if (ext == ".xlsx") and ("月度汇总表" in name):
            tables.append(f)
            dates.append(int(name[-8:]))

    dates = np.array(dates, dtype=np.int32)
    sort_index = np.argsort(dates)
    tables_sorted = []
    for i in sort_index:
        tables_sorted.append(tables[i])

    # read the name list in
    result_df = pd.read_excel("../nameList.xlsx")
    week1 = np.zeros(len(result_df["姓名"]))
    week2 = np.zeros_like(week1)
    week3 = np.zeros_like(week1)
    week4 = np.zeros_like(week1)
    weekList = [week1, week2, week3, week4]
    late = np.zeros_like(week1)

    # gather all the information in the tables
    for f,week in zip(tables_sorted,weekList):
        # read the table in and fill all the NAN, only read in the time of the first check in.
        df = preprocessing(f)
        
        for name, time in zip(df["基本信息"], df["出勤统计"]):
            # collect the working time
            i = result_df.index[result_df["姓名"]==name].tolist()[0]
            week[i] = time
            
            # collecting the number of 'late for work' incidence
            sub_df = df.loc[df["基本信息"]==name]
            late_num = check_late(sub_df)
            late[i] += late_num
            
    result_df["week1"] = week1
    result_df["week2"] = week2
    result_df["week3"] = week3        
    result_df["week4"] = week4
    result_df["late"]  = late
    # result_df["bonus"] = calc_bonus(week1) + calc_bonus(week2) + calc_bonus(week3) + calc_bonus(week4)



    result_df.to_excel("汇总结果-original.xlsx", index=False)        
    