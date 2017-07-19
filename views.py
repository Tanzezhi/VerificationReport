# coding=utf-8

from math import *
from django.shortcuts import render_to_response
from django.views.decorators.csrf import csrf_exempt
import xlrd

data = xlrd.open_workbook(u"./副本室分单验报告模板.xlsx", "rb")
design_value = data.sheet_by_name(u'工参')  # 设计值
measured_value = data.sheet_by_name(u'测试数据CSV')  # 实测值

# 设计值的各项数据
Lontitude = design_value.cell(1, 7).value  # 经度设计值
Latitude = design_value.cell(1, 8).value  # 纬度设计值

CGI = design_value.cell(1, 5).value  # CGI设计值
BCCH = design_value.cell(1, 14).value  # BCCH设计值
BSIC = int(design_value.cell(1, 22).value) + int(design_value.cell(1, 13).value)  # BSIC设计值  BSIC = BCC + NCC
ID = design_value.cell(1, 6).value  # 基站ID设计值
LAC = design_value.cell(1, 3).value  # 归属LAC设计值

CGT_LAC = CGI.split('-')[2]
CGT_CellID = CGI.split('-')[3]

# nrows = design_value.nrows          # 获取行数
# ncols = design_value.ncols          # 获取列数

# 测试值各项数据
measured_value_nrows = measured_value.nrows  # 获取实测表行数
measured_lontitude = 0
measured_latitude = 0
# print measured_value_nrows

"""
判断经纬度是否一致
H4=设计经度，H5=设计纬度，N4=实测经度，N5=实测纬度，判别门限500m
=IF(OR(H4=0,H4=""),"",
IF(INT(6371004*ACOS((SIN(RADIANS(H5))*SIN(RADIANS(N5))+COS(RADIANS(H5))*COS(RADIANS(N5))*COS(RADIANS(N4-H4)))))<500,"是","否"))
"""
for i in range(1, measured_value_nrows):
    measured_lontitude = measured_lontitude + measured_value.cell(i, 8).value  # 经度和
    measured_latitude = measured_latitude + measured_value.cell(i, 7).value  # 纬度和

measured_lontitude_value = measured_lontitude / (measured_value_nrows - 1)  # 经度实测值的平均值
measured_latitude_value = measured_latitude / (measured_value_nrows - 1)  # 纬度实测值的平均值
# print measured_lontitude_value, measured_latitude_value


error_value = int(6371004 * acos((sin(radians(Latitude)) * sin(radians(measured_latitude_value))
                                  + cos(radians(Latitude)) * cos(radians(measured_latitude_value))
                                  * cos(radians(measured_lontitude_value - Lontitude)))))  # 获取误差值判断经纬度是否一致

print error_value

"""
判断CGI和归属LAC是否一致
"""
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 9).value).split('.')[0] == '' \
            or str(measured_value.cell(i, 9).value).split('.')[0] == str(CGT_LAC):
        success_LAC = True
        success_CGI = True
    else:
        success_LAC = False
        success_CGI = False
        break

    if str(measured_value.cell(i, 10).value).split('.')[0] == '' \
            or str(measured_value.cell(i, 10).value).split('.')[0] == str(CGT_CellID):
        success_CGI = True
    else:
        success_CGI = False
        break

print success_CGI, success_LAC

"""
判断BCCH是否一致
"""
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 11).value).split('.')[0] == '' \
            or str(measured_value.cell(i, 11).value).split('.')[0] == str(BCCH).split('.')[0]:
        success_BCCH = True
    else:
        success_BCCH = False
        break
print success_BCCH

"""
判断BSIC是否一致
BSIC=NCC+BCC
BSIC = int(design_value.cell(1, 22).value) + int(design_value.cell(1, 13).value)
"""
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 12).value).split('.')[0] == '' \
            or str(measured_value.cell(i, 12).value).split('.')[0] == str(BSIC).split('.')[0]:
        success_BSIC = True
    else:
        success_BSIC = False
        break
print success_BSIC

"""
判断基站ID是否一致
一致
"""
print ID

"""
获取测试点平均电平（dBm)
"""
RxLevelSub = 0
count = 0
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 15).value) == '':
        continue
    else:
        count = count + 1
        RxLevelSub = RxLevelSub + float(measured_value.cell(i, 15).value)

RxLevelSub_value = RxLevelSub / count

print RxLevelSub_value

"""
获取无线接通率（%）和无线掉话率（%）
"""
count_Call_blocked = 0
count_Call_attempt = 0
count_Call_dropped = 0
count_Call_connected = 0
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 5).value) == "Call blocked":
        count_Call_blocked = count_Call_blocked + 1
    elif str(measured_value.cell(i, 5).value) == "Call attempt":
        count_Call_attempt = count_Call_attempt + 1
    elif str(measured_value.cell(i, 5).value) == "Call dropped":
        count_Call_dropped = count_Call_dropped + 1
    elif str(measured_value.cell(i, 5).value) == "Call connected":
        count_Call_connected = count_Call_connected + 1
print count_Call_blocked, count_Call_attempt, count_Call_dropped, count_Call_connected

connected = 1 - (count_Call_blocked / count_Call_attempt)
unconnected = 1 - (count_Call_dropped / count_Call_connected)

connected_value = "%.2f%%" % (connected * 100)
unconnected_value = "%.2f%%" % (unconnected * 100)
print connected_value, unconnected_value

"""
单用户下行峰值吞吐率（kpbs)
"""
max = 0.00
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 17).value) == '':
        value = 0.00
    else:
        value = float(measured_value.cell(i, 17).value)
    if value >= max:
        max = value

print max

"""
语音RxQuality质量0-4级（%）
"""
count_RxQual = 0
count_lessfour = 0
for i in range(1, measured_value_nrows):
    if str(measured_value.cell(i, 16).value) == '':
        continue
    else:
        count_RxQual = count_RxQual + 1
        if float(measured_value.cell(i, 16).value) < 4:
            count_lessfour = count_lessfour + 1
        else:
            continue
var = count_lessfour / count_RxQual
var_value = "%.2f%%" % (var * 100)
print var_value


def submit(request):
    return render_to_response('submit.html')


@csrf_exempt
def testReport(request):
    if request.method == 'GET':
        return render_to_response('submit.html')
    else:
        excel_file = request.FILES.get('file')
        try:
            f = open("./input_excel_file/" + excel_file.name, "wb")
            for line in excel_file.chunks():
                f.write(line)
            f.close()
        except:
            print "fail"

        return render_to_response('hello.html')
