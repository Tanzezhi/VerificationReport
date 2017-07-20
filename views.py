# coding=utf-8

from math import *
from django.shortcuts import render_to_response
from django.views.decorators.csrf import csrf_exempt
import xlrd
import xlwt
from pyExcelerator import *
from xlutils.copy import copy
from openpyxl import Workbook, load_workbook


def submit(request):
    return render_to_response('submit.html')


@csrf_exempt
def testReport(request):
    if request.method == 'GET':
        return render_to_response('submit.html')
    else:
        excel_file = request.FILES.get('file')
        path = "./input_excel_file/" + excel_file.name
        try:
            f = open(path, "wb")
            for line in excel_file.chunks():
                f.write(line)
            f.close()
        except:
            print "fail"

        data = xlrd.open_workbook(path, "rb")
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
        print measured_lontitude_value, measured_latitude_value


        error_value = int(6371004 * acos((sin(radians(Latitude)) * sin(radians(measured_latitude_value))
                                          + cos(radians(Latitude)) * cos(radians(measured_latitude_value))
                                          * cos(radians(measured_lontitude_value - Lontitude)))))  # 获取误差值判断经纬度是否一致
        if error_value < 500:
            lontitude_latitude_conclusion = "一致"
        else:
            lontitude_latitude_conclusion = "不一致"

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
        if success_LAC and success_CGI:
            CGI_conclusion = "一致"
        else:
            CGI_conclusion = "不一致"
        if success_LAC:
            LAC_conclusion = "一致"
        else:
            LAC_conclusion = "不一致"


        measured_CGI = "460-00-" + str(measured_value.cell(4, 9).value).split('.')[0] + '-' \
                       + str(measured_value.cell(4, 10).value).split('.')[0]
        measured_LAC = str(measured_value.cell(4, 9).value).split('.')[0]
        print success_CGI, success_LAC, measured_CGI

        """
        判断BCCH是否一致
        """
        for i in range(1, measured_value_nrows):
            if str(measured_value.cell(i, 11).value).split('.')[0] == '' \
                    or str(measured_value.cell(i, 11).value).split('.')[0] == str(BCCH).split('.')[0]:
                BCCH_conclusion = "一致"
            else:
                BCCH_conclusion = "不一致"
                break
        measured_BCCH = str(measured_value.cell(3, 11).value).split('.')[0]


        """
        判断BSIC是否一致
        BSIC=NCC+BCC
        BSIC = int(design_value.cell(1, 22).value) + int(design_value.cell(1, 13).value)
        """
        for i in range(1, measured_value_nrows):
            if str(measured_value.cell(i, 12).value).split('.')[0] == '' \
                    or str(measured_value.cell(i, 12).value).split('.')[0] == str(BSIC).split('.')[0]:
                BSIC_conclusion = "一致"
            else:
                BSIC_conclusion = "不一致"
                break

        measured_BSIC = str(measured_value.cell(3, 12).value).split('.')[0]

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



        wr = Write_excel(filename=path)
        wr.write('E16', Lontitude)                                       # 填写经度设计值
        wr.write('E17', Latitude)                                        # 填写纬度设计值

        wr.write('H16', measured_lontitude_value)                        # 填写经度实测值
        wr.write('H17', measured_latitude_value)                         # 填写纬度实测值

        wr.write('K16', lontitude_latitude_conclusion)                   # 填写判断经纬度是否一致
        wr.write('K17', lontitude_latitude_conclusion)

        wr.write('E22', CGI)                                             # 填写CGI设计值
        wr.write('E23', BCCH)                                            # 填写BCCH设计值
        wr.write('E24', BSIC)                                            # 填写BSIC设计值
        wr.write('E25', ID)                                              # 填写基站ID设计值
        wr.write('E26', LAC)                                             # 填写归属LAC设计值

        wr.write('H22', measured_CGI)                                    # 填写实测值
        wr.write('H23', measured_BCCH)
        wr.write('H24', measured_BSIC)
        wr.write('H25', ID)
        wr.write('H26', measured_LAC)

        wr.write('K22', CGI_conclusion)
        wr.write('K23', BCCH_conclusion)
        wr.write('K24', BSIC_conclusion)
        wr.write('K25', "一致")
        wr.write('K26', LAC_conclusion)

        wr.write('E36', RxLevelSub_value)                       # 填写测试点平均电平
        wr.write('E37', connected_value)                        # 填写无线接通率
        wr.write('E38', unconnected_value)                      # 填写无线掉话率
        wr.write('E39', max)                                    # 填写单用户下行峰值吞吐率（kpbs)
        wr.write('E40', var_value)                              # 填写语音RxQuality质量0-4级（%）

        return render_to_response('hello.html')


# excel表写操作

class Write_excel(object):
    def __init__(self, filename):
        self.filename = filename
        self.wb = load_workbook(self.filename)
        self.ws = self.wb.active

    def write(self, coord, value):
        # eg: coord:A1
        self.ws.cell(coord).value = value
        self.wb.save(self.filename)



