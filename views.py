# coding=utf-8
from math import *
from django.shortcuts import render_to_response
from django.views.decorators.csrf import csrf_exempt
import xlrd
import xlwt
from pyExcelerator import *
from xlutils.copy import copy
from openpyxl import Workbook, load_workbook
import sys
from collections import Counter



"""
获取经纬度的值并判断经纬度是否一致
H4=设计经度，H5=设计纬度，N4=实测经度，N5=实测纬度，判别门限500m
=IF(OR(H4=0,H4=""),"",
IF(INT(6371004*ACOS((SIN(RADIANS(H5))*SIN(RADIANS(N5))+COS(RADIANS(H5))*COS(RADIANS(N5))*COS(RADIANS(N4-H4)))))<500,"是","否"))
"""


def get_lontitude_latitude(Latitude, Lontitude, measured_sheet, measured_sheet_nrows):
    measured_lontitude = 0
    measured_latitude = 0
    for i in range(1, measured_sheet_nrows):
        measured_lontitude = measured_lontitude + measured_sheet.cell(i, 8).value  # 经度和
        measured_latitude = measured_latitude + measured_sheet.cell(i, 7).value  # 纬度和

    measured_lontitude_value = measured_lontitude / (measured_sheet_nrows - 1)  # 经度实测值的平均值
    measured_latitude_value = measured_latitude / (measured_sheet_nrows - 1)  # 纬度实测值的平均值

    error_value = int(6371004 * acos((sin(radians(Latitude)) * sin(radians(measured_latitude_value))
                                      + cos(radians(Latitude)) * cos(radians(measured_latitude_value))
                                      * cos(radians(measured_lontitude_value - Lontitude)))))  # 获取误差值判断经纬度是否一致
    if error_value < 500:
        lontitude_latitude_conclusion = "一致"
    else:
        lontitude_latitude_conclusion = "不一致"

    return measured_lontitude_value, measured_latitude_value, lontitude_latitude_conclusion, error_value


"""
获取CGI和LAC的值并判断CGI和归属LAC是否一致
"""


def get_CGI_LAC(CGT_LAC, CGT_CellID, measured_sheet, measured_sheet_nrows):
    for i in range(1, measured_sheet.nrows):
        if str(measured_sheet.cell(i, 9).value).split('.')[0] != str(CGT_LAC):
            success_LAC = False
        else:
            success_LAC = True
            break
    for i in range(1, measured_sheet.nrows):
        if str(measured_sheet.cell(i, 10).value).split('.')[0] != str(CGT_CellID):
            success_CellID = False
        else:
            success_CellID = True
            break
    if success_LAC and success_CellID:
        CGI_conclusion = "一致"
    else:
        CGI_conclusion = "不一致"
    if success_LAC:
        LAC_conclusion = "一致"
    else:
        LAC_conclusion = "不一致"

    LAC = []  # 所有的LAC字符串
    for i in range(1, measured_sheet.nrows):
        LAC.append(str(measured_sheet.cell(i, 9).value).split('.')[0])
    LAC_list = Counter(LAC).most_common(1)  # [()]
    LAC_most = (LAC_list[0])[0]  # 得到出现次数最多的LAC值

    CellID = []  # 所有的CellID字符串
    for i in range(1, measured_sheet.nrows):
        CellID.append(str(measured_sheet.cell(i, 10).value).split('.')[0])
    CellID_list = Counter(CellID).most_common(1)
    CellID_most = (CellID_list[0])[0]  # 得到出现次数最多的CellID值

    measured_CGI = "460-00-" + LAC_most + '-' \
                   + CellID_most
    measured_LAC = LAC_most

    return CGI_conclusion, LAC_conclusion, measured_CGI, measured_LAC


"""
获取BCCH的值并判断BCCH是否一致
"""


def get_BCCH(BCCH, measured_sheet, measured_sheet_nrows):
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 11).value).split('.')[0] != str(BCCH).split('.')[0]:
            BCCH_conclusion = "不一致"
        else:
            BCCH_conclusion = "一致"
            break

    list = []  # 所有的BCCH字符串
    for i in range(1, measured_sheet.nrows):
        list.append(str(measured_sheet.cell(i, 11).value).split('.')[0])
    BCCH_list = Counter(list).most_common(1)
    BCCH_most = (BCCH_list[0])[0]  # 得到出现次数最多的CellID值
    return BCCH_conclusion, BCCH_most


"""
判断BSIC是否一致
BSIC=NCC+BCC
BSIC = int(design_value.cell(1, 22).value) + int(design_value.cell(1, 13).value)
"""


def get_BSIC(BSIC, measured_sheet, measured_sheet_nrows):
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 12).value).split('.')[0] != str(BSIC).split('.')[0]:
            BSIC_conclusion = "不一致"
        else:
            BSIC_conclusion = "一致"

    list = []  # 所有的BCCH字符串
    for i in range(1, measured_sheet.nrows):
        list.append(str(measured_sheet.cell(i, 12).value).split('.')[0])
    BSIC_list = Counter(list).most_common(1)
    BSIC_most = (BSIC_list[0])[0]  # 得到出现次数最多的CellID值

    return BSIC_conclusion, BSIC_most


"""
判断基站ID是否一致
一致
"""


def get_ID(id, measured_sheet, measured_sheet_nrows):
    list = []  # 所有的CellID字符串
    for i in range(1, measured_sheet.nrows):
        list.append(str(measured_sheet.cell(i, 10).value).split('.')[0])
    ID_list = Counter(list).most_common(1)
    ID_most = (ID_list[0])[0]  # 得到出现次数最多的CellID值
    ID_bits = ('%x' % int(ID_most)).zfill(4)
    measured_ID = id + str(ID_bits).upper()
    return measured_ID


"""
获取测试点平均电平（dBm)
"""


def get_RxLevelSub_value(measured_sheet, measured_sheet_nrows):
    RxLevelSub = 0
    count = 0
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 15).value) == '':
            continue
        else:
            count = count + 1
            RxLevelSub = RxLevelSub + float(measured_sheet.cell(i, 15).value)

    RxLevelSub_value = RxLevelSub / count

    return RxLevelSub_value


"""
获取无线接通率（%）和无线掉话率（%）
"""


def get_connected_value_and_unconnected_value(measured_sheet, measured_sheet_nrows):
    count_Call_blocked = 0
    count_Call_attempt = 0
    count_Call_dropped = 0
    count_Call_connected = 0
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 5).value) == "Call blocked":
            count_Call_blocked = count_Call_blocked + 1
        elif str(measured_sheet.cell(i, 5).value) == "Call attempt":
            count_Call_attempt = count_Call_attempt + 1
        elif str(measured_sheet.cell(i, 5).value) == "Call dropped":
            count_Call_dropped = count_Call_dropped + 1
        elif str(measured_sheet.cell(i, 5).value) == "Call connected":
            count_Call_connected = count_Call_connected + 1

    connected = 1 - (count_Call_blocked / count_Call_attempt)
    unconnected = 1 - (count_Call_dropped / count_Call_connected)

    connected_value = "%.2f%%" % (connected * 100)
    unconnected_value = "%.2f%%" % (unconnected * 100)
    return connected_value, unconnected_value


"""
单用户下行峰值吞吐率（kpbs)
"""


def get_max(measured_sheet, measured_sheet_nrows):
    max = 0.00
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 17).value) == '':
            value = 0.00
        else:
            value = float(measured_sheet.cell(i, 17).value)
        if value >= max:
            max = value

    return max


"""
语音RxQuality质量0-4级（%）
"""


def var_value(measured_sheet, measured_sheet_nrows):
    count_RxQual = 0
    count_lessfour = 0
    for i in range(1, measured_sheet_nrows):
        if str(measured_sheet.cell(i, 16).value) == '':
            continue
        else:
            count_RxQual = count_RxQual + 1
            if float(measured_sheet.cell(i, 16).value) < 4:
                count_lessfour = count_lessfour + 1
            else:
                continue
    var = count_lessfour / count_RxQual
    value = "%.2f%%" % (var * 100)
    return value


# 读取excel存放在input_excel_file文件夹下

def input_excel_file(path, file_name):
    try:
        f = open(path, "wb")
        for line in file_name.chunks():
            f.write(line)
        f.close()
    except:
        print "fail"



def submit(request):
    return render_to_response("submit_1.html")


@csrf_exempt
def testReport(request):
    reload(sys)
    sys.setdefaultencoding('utf-8')


    template_file = request.FILES.get('template_file')                      # 从客户端获取模板excel
    template_path = "./input_excel_file/" + template_file.name              # 模板excel存储路径
    input_excel_file(template_path, template_file)                          # 读取模板excel存放在input_excel_file文件夹下

    test_1_file = request.FILES.get('s1_file')                              # 从客户端获取测试excel
    test_1_path = "./input_excel_file/" + test_1_file.name                  # 测试excel存储路径
    input_excel_file(test_1_path, test_1_file)                              # 读取测试excel存放在input_excel_file文件夹下

    village_name = request.POST['s1_name']                                  # 从客户端获取小区名用于在模板excel中查询相应小区信息




    # 副本室分单验报告模板文件

    data = xlrd.open_workbook(template_path, "rb")
    design_sheet = data.sheet_by_name(u'工参')  # 设计值
    design_sheet_nrows = design_sheet.nrows



    # 测试数据文件

    measured_data = xlrd.open_workbook(test_1_path, "rb")
    measured_sheet = measured_data.sheet_by_name(u'测试数据CSV')
    measured_sheet_nrows = measured_sheet.nrows




    for i in range(1, design_sheet_nrows):
        if str(design_sheet.cell(i, 1).value) != village_name:
            success = False
        else:
            success = True
            break

    if success:
        village_row = i

        design_value = Design_sheet(village_row, design_sheet)  # 获取所要的设计数据

        # 经纬度的实测值和是否一致
        measured_lontitude_value, measured_latitude_value, lontitude_latitude_conclusion, error_value = \
            get_lontitude_latitude(design_value.Latitude, design_value.Lontitude, measured_sheet, measured_sheet_nrows)
        print measured_lontitude_value, measured_latitude_value, lontitude_latitude_conclusion, error_value

        # CGI和LAC实测值和是否一致
        CGI = design_value.CGI
        CGT_LAC = CGI.split('-')[2]
        CGT_CellID = CGI.split('-')[3]
        CGI_conclusion, LAC_conclusion, measured_CGI, measured_LAC = get_CGI_LAC(CGT_LAC, CGT_CellID, measured_sheet, measured_sheet_nrows)
        print CGI_conclusion, LAC_conclusion, measured_CGI, measured_LAC

        # BCCH实测值和是否一致
        BCCH_conclusion, BCCH_most = get_BCCH(design_value.BCCH, measured_sheet, measured_sheet_nrows)
        print BCCH_conclusion, BCCH_most

        # BSIC的实测值和是否一致
        BSIC_conclusion, BSIC_most = get_BSIC(design_value.BSIC, measured_sheet, measured_sheet_nrows)
        print BSIC_conclusion, BSIC_most

        # 获取测试点平均电平（dBm)
        RxLevelSub_value = get_RxLevelSub_value(measured_sheet, measured_sheet_nrows)
        print RxLevelSub_value

        # 获取无线接通率（%）和无线掉话率（%）
        connected_value, unconnected_value = get_connected_value_and_unconnected_value(measured_sheet, measured_sheet_nrows)
        print connected_value, unconnected_value

        # 单用户下行峰值吞吐率（kpbs)
        max = get_max(measured_sheet, measured_sheet_nrows)
        print max

        # 语音RxQuality质量0-4级（%）
        value = var_value(measured_sheet, measured_sheet_nrows)
        print value

        # 判断基站ID是否一致
        id = design_value.MSC.split('-')[1]
        measured_ID = get_ID(id, measured_sheet, measured_sheet_nrows)
        print measured_ID
        print id


        wr = Write_excel(filename=template_path)
        wr.write('E16', design_value.Lontitude)  # 填写经度设计值
        wr.write('E17', design_value.Lontitude)  # 填写纬度设计值

        wr.write('H16', measured_lontitude_value)  # 填写经度实测值
        wr.write('H17', measured_latitude_value)  # 填写纬度实测值

        wr.write('K16', lontitude_latitude_conclusion)  # 填写判断经纬度是否一致
        wr.write('K17', lontitude_latitude_conclusion)

        wr.write('E22', CGI)  # 填写CGI设计值
        wr.write('E23', design_value.BCCH)  # 填写BCCH设计值
        wr.write('E24', design_value.BSIC)  # 填写BSIC设计值
        wr.write('E25', design_value.ID)  # 填写基站ID设计值
        wr.write('E26', design_value.LAC)  # 填写归属LAC设计值

        wr.write('H22', measured_CGI)  # 填写实测值
        wr.write('H23', BCCH_most)
        wr.write('H24', BSIC_most)
        wr.write('H25', measured_ID)
        wr.write('H26', measured_LAC)

        wr.write('K22', CGI_conclusion)
        wr.write('K23', BCCH_conclusion)
        wr.write('K24', BSIC_conclusion)
        wr.write('K25', "一致")
        wr.write('K26', LAC_conclusion)

        wr.write('E36', RxLevelSub_value)  # 填写测试点平均电平
        wr.write('E37', connected_value)  # 填写无线接通率
        wr.write('E38', unconnected_value)  # 填写无线掉话率
        wr.write('E39', max)  # 填写单用户下行峰值吞吐率（kpbs)
        wr.write('E40', value)  # 填写语音RxQuality质量0-4级（%）

        return render_to_response("hello.html")
    else:
        return render_to_response("not_find.html")


class Design_sheet(object):
    def __init__(self, village_row, design_sheet):
        self.Lontitude = float(design_sheet.cell(village_row, 7).value)
        self.Latitude = float(design_sheet.cell(village_row, 8).value)
        self.CGI = str(design_sheet.cell(village_row, 5).value)
        self.BCCH = str(design_sheet.cell(village_row, 14).value).split('.')[0]
        self.BSIC = int(design_sheet.cell(village_row, 22).value) * 10 + int(design_sheet.cell(village_row, 13).value)
        self.ID = str(design_sheet.cell(village_row, 6).value)
        self.LAC = str(design_sheet.cell(village_row, 3).value)
        self.MSC = str(design_sheet.cell(village_row, 31).value)


class Write_excel(object):
    def __init__(self, filename):
        self.filename = filename
        self.wb = load_workbook(self.filename)

        self.ws = self.wb.active

    def write(self, coord, value):
        # eg: coord:A1
        self.ws.cell(coord).value = value
        self.wb.save(self.filename)
