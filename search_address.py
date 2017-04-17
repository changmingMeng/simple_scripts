# encoding: utf-8

import os
import re
import xlrd
import psycopg2
import datetime
import logging

import utils

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename='address.log',
                    filemode='w')

console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)


no_gz_citys=["深圳", "珠海", "汕头", "韶关", "佛山", "江门",
                "湛江", "茂名", "肇庆", "惠州", "梅州", "汕尾",
                "河源", "阳江", "清远", "东莞", "中山", "潮州",
                "揭阳", "云浮"]
district_of_gz = ["天河", "海珠", "越秀", "白云", "番禺", "荔湾",
                  "黄埔", "从化", "花都", "增城", "南沙",
                  "萝岗", "黄浦"]
dict_gz = {district:0 for district in district_of_gz}
def get_sequence_num_of_letter(letter):
    #print letter
    return ord(letter.upper())-ord('A')+1

def get_excel_col_number(str):
    if len(str) == 1:
        return get_sequence_num_of_letter(str[0])-1
    elif len(str) == 2:
        return get_sequence_num_of_letter(str[0])*26 +get_sequence_num_of_letter(str[1])-1
    else:
        raise("too many letters")

def search_address(filename, sheet_num, type_col_letter,address_col_letter):
    with xlrd.open_workbook(filename) as workbook:
        sheet = workbook.sheet_by_index(sheet_num)

    tianhe = 0
    haizhu = 0
    yuexiu = 0
    liwan = 0
    panyu = 0
    baiyun = 0
    huadu = 0
    huangpu = 0
    conghua = 0
    zengcheng = 0
    nansha = 0
    not_gz = 0

    row_number = 1
    net_error = 0
    for r in xrange(1,sheet.nrows):#excel第2行到倒数第2行
        row = sheet.row_values(r)
        row_number += 1
        print row[get_excel_col_number(type_col_letter)]
        if row[get_excel_col_number(type_col_letter)] == '广州'.decode('utf-8'):
            net_error += 1
            #print row[78], row[79], row[80]
            address = row[get_excel_col_number(address_col_letter)].encode('utf-8')

            if re.search('天河', address):
                tianhe += 1
                #logging.debug('天河'+str(row_number) + address)
            elif re.search('海珠', address):
                haizhu += 1
                #logging.debug('海珠' + str(row_number) + address)
            elif re.search('越秀', address):
                yuexiu += 1
                #logging.debug('越秀' + str(row_number) + address)
            elif re.search('荔湾', address):
                liwan += 1
                #logging.debug('荔湾' + str(row_number) + address)
            elif re.search('番禺', address):
                panyu += 1
                #logging.debug('番禺' + str(row_number) + address)
            elif re.search('白云', address):
                baiyun += 1
                #logging.debug('白云' + str(row_number) + address)
            elif re.search('花都', address):
                huadu += 1
                #logging.debug('花都' + str(row_number) + address)
            elif re.search('黄埔', address) or re.search('萝岗', address) or re.search('黄浦', address):
                huangpu += 1
                #logging.debug('黄埔' + str(row_number) + address)
            elif re.search('从化', address):
                conghua += 1
                #logging.debug('从化' + str(row_number) + address)
            elif re.search('增城', address):
                zengcheng += 1
                #logging.debug('增城' + str(row_number) + address)
            elif re.search('南沙', address):
                nansha += 1
                #logging.debug('南沙' + str(row_number) + address)
            else:
                for city in no_gz_citys:
                    if re.search(city, address):
                        logging.debug(str(row_number)+"非广州")
                        not_gz += 1
                # logging.debug(str(row_number)+address)

    logging.info('天河：'+str(tianhe))
    logging.info('海珠：' + str(haizhu))
    logging.info('越秀：' + str(yuexiu))
    logging.info('荔湾：' + str(liwan))
    logging.info('番禺：' + str(panyu))
    logging.info('白云：' + str(baiyun))
    logging.info('花都：' + str(huadu))
    logging.info('黄埔：' + str(huangpu))
    logging.info('从化：' + str(conghua))
    logging.info('增城：' + str(zengcheng))
    logging.info('南沙：' + str(nansha))
    logging.info('非广州：' + str(not_gz))

    logging.info('总共：' + str(row_number-1))
    logging.info('网络原因：' + str(net_error))
    logging.info('识别：' + str(tianhe+haizhu+yuexiu+panyu+liwan+baiyun+huadu+huangpu+conghua+zengcheng+nansha))

def analysis(description):
    description = description.encode('utf-8')
    print description
    for city in no_gz_citys:
        if re.search(city, description):
            address_str = description[description.find(city):]
            return ["否", "非广州", address_str]

    for district in dict_gz.keys():
        if re.search(district, description):
            dict_gz[district] += 1
            address_str = description[description.find(district):]#description.find(district)+200
            return ["是", district, address_str]

    return ["未识别","",description]
    # if re.search("村", description):
    #     return ["是", "覆盖问题", "城中村"]


def complain_analysis(filename, outputfilename, sheet_num, reason_col_letter, anchor_col_letter):
    with xlrd.open_workbook(filename) as workbook:
        sheet = workbook.sheet_by_index(sheet_num)

    row_number = 1
    lst_lst_result = []
    for r in xrange(1, sheet.nrows):  # excel第2行到倒数第2行
        row = sheet.row_values(r)
        row_number += 1
        description = row[get_excel_col_number(reason_col_letter)]
        lst_result = analysis(description)

        is_net, district, describe = lst_result
        print str(row_number), is_net, district, describe

        lst_result.append(row[get_excel_col_number(anchor_col_letter)])
        lst_lst_result.append(lst_result)

        #logging.INFO(str(row_number) + " is_net:" + is_net + " 区域：" + district + " 描述：" + describe)

    lst_name = ["是否投诉网络质量问题", "区域", "描述", "锚定"]
    utils.write_lst_of_lst_to_excel(outputfilename, lst_name, lst_lst_result)

if __name__ == "__main__":
    #search_address(r'F:\视频组\投诉\4-13\副本网络详单0405-0411.xlsx'.decode('utf-8').encode('GBK'), 0, "A", "T")
    #print get_excel_col_number('CJ')
    # str1="AB"
    # print str1[0]
    complain_analysis(r'F:\视频组\投诉\4-13\副本网络详单0405-0411.xlsx'.decode('utf-8').encode('GBK'),
                      r'F:\视频组\投诉\4-13\副本网络详单0405-0411_result.xlsx'.decode('utf-8').encode('GBK'),
                      0,
                      "X",
                      "Q")

