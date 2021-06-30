# man_hour_collector
用法：
1.创建data文件夹，把数据都放在data下，其中数据首行必须是表头
2.运行collect_data.py来把表格整合到一个文件内，输出raw_data.xlsx
3.运行check_element.py查找低级错误，包括空行、日期超出范围，输出summary.xlsx
4.运行check_man_hour.py统计工时，输出check_man_hour.xlsx
