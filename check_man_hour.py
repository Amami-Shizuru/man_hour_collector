import datetime
import pandas as pd

def checkManHours(raw_diary, raw_weekly):
    for index, row in raw_weekly.iterrows():
        actual_man_hour = row['本周实际工时']
        sum = 0
        for daily_index, daily_row in raw_diary.iterrows():
            if row['姓名'] == daily_row['姓名'] and row['项目代码'] == daily_row['项目代码']:
                sum += daily_row['工时']
        raw_weekly.iloc[index, -1] = sum
        if actual_man_hour != sum:
            raw_weekly.iloc[index, -2] = '每日工时总和不等于总工时'

    return raw_weekly

if __name__ == '__main__':
    data = pd.read_excel('./summary.xlsx', sheet_name=['本周工作日志', '周总结及计划'], header=0, keep_default_na=False)
    raw_diary = data['本周工作日志']
    raw_weekly = data['周总结及计划']
    raw_weekly['异常'] = ''
    raw_weekly['每日工时总和'] = ''
    weekly = checkManHours(raw_diary, raw_weekly)

    writer = pd.ExcelWriter('./man_hour_check.xlsx')
    raw_diary.to_excel(writer, index=False, sheet_name='本周工作日志')
    weekly.to_excel(writer, index=False, sheet_name='周总结及计划')
    writer.save()
