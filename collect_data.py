import os
import pandas as pd

def stripSpace(raw_data):
    for index, row in raw_data.iterrows():
        for key, value in row.items():
            # strip redundant spaces
            if isinstance(value, str):
                row[key] = value.strip()
        raw_data.iloc[index, :] = row
    return raw_data

def collectDiaryAndWeekly(dir):
    diary_list = []
    weekly_list = []
    for root, dirs, files in os.walk(dir):
        for f in files:
            if f.endswith(".xlsx"):
                file_path = os.path.join(root, f)
                print("Dumping data from " + file_path)
                raw_data = pd.read_excel(file_path, sheet_name=['本周工作日志','周总结及计划'], header=0, keep_default_na=False)
                raw_diary = raw_data['本周工作日志'].loc[:, ~raw_data['本周工作日志'].columns.str.contains('Unnamed')]
                raw_weekly = raw_data['周总结及计划'].loc[:, ~raw_data['周总结及计划'].columns.str.contains('Unnamed')]
                diary_list.append(raw_diary)
                weekly_list.append(raw_weekly)
                print("Finish dumping data from " + file_path)
    diary = pd.concat(diary_list, axis=0)
    weekly = pd.concat(weekly_list, axis=0)
    diary.index = [i for i in range(len(diary))]
    weekly.index = [i for i in range(len(weekly))]
    print("Dump data finished!")
    print("Stripping spaces in dataframe")
    diary = stripSpace(diary)
    weekly = stripSpace(weekly)
    print("Strip spaces success!")
    return diary, weekly


if __name__ == '__main__':
    diary, weekly = collectDiaryAndWeekly('./data')
    print("Saving to raw_data.xlsx")
    writer = pd.ExcelWriter('./raw_data.xlsx')
    diary.to_excel(writer, index=False, sheet_name='本周工作日志')
    weekly.to_excel(writer, index=False, sheet_name='周总结及计划')
    writer.save()
