# -*- coding: utf-8 -*-
# Author: Jason Ke

import xlrd
import xlwt
from xlutils.copy import copy

class ExcelAnalysis():
    originDataFile = './原始数据.xlsx'
    parseDataFile = './数据解析.xlsx'
    # outputPath = './输出文件夹/'
    fileName = 'results'

    # 获取文件的位置
    def __get_file_name(self):
        originFileName = input('请输入原始数据文件名：(回车跳过默认为“原始数据.xlsx”)：')
        parseFileName = input('请输入解析数据文件名：(回车跳过默认为“数据解析.xlsx”)：')
        # outputDocName = input('请输入存放结果文件夹名称：(回车跳过默认为“输出文件夹”)：')
        resultFileName = input('请输入保存时Excel文件名称：(回车跳过默认为“results”)：')
        if originFileName != '':
            ExcelAnalysis.originDataFile = originFileName if '.xlsx' in originFileName else originFileName + '.xlsx'
        if parseFileName != '':
            ExcelAnalysis.parseDataFile = parseFileName if '.xlsx' in parseFileName else parseFileName + '.xlsx'
        # if outputDocName != '':
        #     ExcelAnalysis.outputPath = outputDocName if '.xlsx' in outputDocName else './' + outputDocName + '/'
        if resultFileName != '':
            ExcelAnalysis.fileName = resultFileName if '.xlsx' in resultFileName else resultFileName

    # 解析原始数据成数组
    def __get_origin_data(self):
        originData = self.__get_excel_data(ExcelAnalysis.originDataFile)
        return originData

    # 解析匹配数据成数组
    def __get_parse_data(self):
        parseData= self.__get_excel_data(ExcelAnalysis.parseDataFile)
        afterTranslateData = []
        for pd in parseData:
            record = []
            companyName = pd[0]
            parseWords = pd[1].split('$')
            record.append(companyName)
            record.append(parseWords)
            afterTranslateData.append(record)
        return afterTranslateData
    
    # 工具方法
    # 读Excel
    def __get_excel_data(self, fileName):
        excel = xlrd.open_workbook(fileName)
        sheet = excel.sheets()[0]
        data = []
        # 如果当前Excel含有标题，从第1行开始读取，而非第0行
        for i in range(1, sheet.nrows):
            rowData = []
            for j in range(sheet.ncols):
                rowData.append(sheet.cell_value(i, j))
            data.append(rowData)
        return data

    # 设置Excel样式
    def __excel_set_style(self, name, height, blod = False):
        style = xlwt.XFStyle()

        font = xlwt.Font()
        font.name = name
        font.height = height
        font.color_index = 4
        font.blod = blod
        style.font = font

        return style
    # 写Excel
    def __write_excel_data(self, dataArray):
        # 当前数据中任意取一条数据的第一条数据当做文件名
        # newFileName = dataArray[0][0]
        newFileName = ExcelAnalysis.fileName
        f = xlwt.Workbook()
        sheet = f.add_sheet(u'sheet1', cell_overwrite_ok=False)

        # 写入Title
        row0 = ['公司名称']
        for i in range(0, len(row0)):
            sheet.write(0, i, row0[i], self.__excel_set_style('Times New Roman', 220, True))
        
        # 写入正文
        for line in range(len(dataArray)):
            for cell in range(len(dataArray[line])):
                sheet.write(line+1, cell, dataArray[line][cell], self.__excel_set_style('Arial', 200, True))
        # f.save(ExcelAnalysis.outputPath + newFileName + '.xlsx')
        f.save(newFileName + '.xlsx')

    # 处理业务逻辑
    def __handle_business(self, originData, parseData):
        result = []
        for pd in parseData:
            for od in originData:
                rowResult = []
                isContains = True
                index = 1
                for keyWord in pd[1]:
                    if not(keyWord in str(od)):
                        isContains = False
                        break
                    index += 1
                if isContains:
                    rowResult.append(pd[0])
                    rowResult.extend(od)
                if len(rowResult) > 0:
                    result.append(rowResult)
        return result
    '''
    将所有结果根据公司名称写入不同的Excel
    暂时不用该方法
    TODO
    def __write_to_diff_excel(self, result):
        for r in result:
            self.__write_excel_data(r)
    '''

    def go(self):
        self.__get_file_name()
        # 原始数据
        originData = self.__get_origin_data()
        parseData = self.__get_parse_data()
        parseResult = self.__handle_business(originData, parseData)
        self.__write_excel_data(parseResult)
        
ea = ExcelAnalysis()
ea.go()