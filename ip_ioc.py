#!/usr/bin/python3
# -*- coding: utf-8 -*-
import argparse
import json
import linecache
import xlwt


# python3环境运行
# {"ip_value": "101.0.0.1", "response_code": 0, "result": "无数据"}
# {"ip_value": "101.0.0.0", "response_code": 1, "tags": ["Malicious Host", "malicious", "samba", "可疑"], "ip_address": "中国上海市上海市", "source": "开源情报", "time": "2018-12-12 00:00:00"}
# 参数 -f 上面两行内容格式相同的文件


resulsts_list = []


# 初始化数据
def init_data(filename):
    file_lines = linecache.getlines(filename)
    for line in file_lines:
        if len(line) > 10:
            try:
                line_json = json.loads(line)
                resulsts_list.append(line_json)
            except Exception as e:
                print(e)
        else:
            pass


def list_to_str(data_list):
    data_str = ""
    for data in data_list:
        data_str = data_str + data + ','
    data_str = data_str[0:len(data_str)-1]
    return data_str


# 数据保存到excel
def save_data_to_excel(excel_filename):
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('sheet1')
    worksheet.write(0, 0, 'ip')
    worksheet.write(0, 1, 'result')
    worksheet.write(0, 2, '地区')
    count = 0
    for result in resulsts_list:
        count += 1
        if result["response_code"] == 0:
            worksheet.write(count, 0, result["ip_value"])
            worksheet.write(count, 1, result["result"])
        elif result["response_code"] == 1:
            worksheet.write(count, 0, result["ip_value"])
            worksheet.write(count, 1, list_to_str(result["tags"]))
            try:
                if "ip_address" in result:
                    if result["ip_address"] != 'null':
                        worksheet.write(count, 2, result["ip_address"])
                    else:
                        pass
                else:
                    pass
            except Exception as e:
                print(count)
                print(e)
        else:
            pass
    workbook.save(excel_filename)


if __name__ == '__main__':
    p = argparse.ArgumentParser(description='Port scanner!.')
    p.add_argument('-f', dest='file_name', type=str)
    args = p.parse_args()
    file_name = args.file_name
    init_data(file_name)
    excel_name = file_name + '-results.xls'
    save_data_to_excel(excel_name)
