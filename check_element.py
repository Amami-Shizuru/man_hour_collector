import datetime
import pandas as pd


def addException(str, exception):
    if str == '':
        return exception
    else:
        str = str + '，' + exception
        return str


# get begin and end date
def getBeginAndEndDate():
    input_begin = input("请输入起始日期，格式：yyyy-mm-dd：\n").split("-")
    assert len(input_begin) == 3, '日期格式错误'
    date_begin = datetime.date(year=int(input_begin[0]), month=int(input_begin[1]), \
                               day=int(input_begin[2]))

    input_end = input("请输入结束日期，格式：yyyy-mm-dd：\n").split("-")
    assert len(input_end) == 3, '日期格式错误'
    date_end = datetime.date(year=int(input_end[0]), month=int(input_end[1]), \
                             day=int(input_end[2]))
    assert date_begin <= date_end, '起始日期晚于结束日期'
    return date_begin, date_end


def checkDiaryElement(raw_data, date_begin, date_end):
    # append exception column
    raw_data['异常'] = ''

    # set valid void column
    valid_nan = ['阶段', '任务名称', '异常']

    for index, row in raw_data.iterrows():
        for key, value in row.items():
            # strip redundant spaces
            if isinstance(value, str):
                row[key] = value.strip()

            # check invalid void element
            if key not in valid_nan and value == '':
                if set(row.values) == {''}:
                    raw_data = raw_data.drop(index=index)
                    break
                else:
                    row[key] = addException(row[key], '异常空值')
                    break

            # check invalid date
            if key == '日期':
                raw_date = str(int(value))
                year = int(raw_date[:4])
                month = int(raw_date[4:6])
                day = int(raw_date[6:])
                date = datetime.date(year=year, month=month, day=day)
                if date < date_begin or date > date_end:
                    row[key] = addException(row[key], '日期超出范围')

            raw_data.iloc[index,:] = row

    return raw_data


def checkWeeklyElement(raw_weekly):
    # set valid void column
    raw_weekly['异常'] = ''
    invalid_nan = ['姓名', '项目代码', '本周实际工时']

    for index, row in raw_weekly.iterrows():
        for key, value in row.items():
            # strip redundant spaces
            if isinstance(value, str):
                row[key] = value.strip()

            # check invalid void element
            if key in invalid_nan and value == '':
                if set(row.values) == {''}:
                    raw_data = raw_data.drop(index=index)
                    break
                else:
                    raw_data.iloc[index, -1] = addException(raw_data.iloc[index, -1], '异常空值')
                    break
    return raw_weekly


if __name__ == '__main__':
    raw_data = pd.read_excel('./raw_data.xlsx', sheet_name=['本周工作日志', '周总结及计划'], header=0, keep_default_na=False)
    raw_diary = raw_data['本周工作日志']
    raw_weekly = raw_data['周总结及计划']

    date_begin, date_end = getBeginAndEndDate()
    diary = checkDiaryElement(raw_diary, date_begin, date_end)
    weekly = checkWeeklyElement(raw_weekly)

    writer = pd.ExcelWriter('./summary.xlsx')
    diary.to_excel(writer, index=False, sheet_name='本周工作日志')
    weekly.to_excel(writer, index=False, sheet_name='周总结及计划')
    writer.save()
