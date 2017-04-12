#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlrd,xlwt
import csv
import os

fo=open('/home/zsf/下载/结果.csv','wb')
writer = csv.writer(fo)
writer.writerow(['文件名', '营业收入', '营业费用','债务','退款','FBA费用','应收款'])

pathDir = os.listdir('/home/zsf/下载/表格/')
for allDir in pathDir:
    child = os.path.join('%s%s' % ('/home/zsf/下载/表格/', allDir))
    if child[-3:]=='xls' or child[-3:]=='lsx':
        data = xlrd.open_workbook(child)
        table = data.sheets()[0]
        income=0
        fee=0
        debt=0
        refund=0
        fbafee=0
        sum=0
        transfer=0
        for i in range(table.nrows):
            if table.row_values(i)[2]=='Order':
                income=income+table.row_values(i)[12]+table.row_values(i)[13]+table.row_values(i)[14]+table.row_values(i)[15]
                fee=fee+table.row_values(i)[16]+table.row_values(i)[17]+table.row_values(i)[18]+table.row_values(i)[19]+table.row_values(i)[20]
            elif table.row_values(i)[2]=='Refund' or table.row_values(i)[2]=='Ajustment' or table.row_values(i)[2]=='ATOZ':
                refund=refund+table.row_values(i)[21]
            elif table.row_values(i)[2]=='FBA Inventory Fee' or table.row_values(i)[2]=='Service Fee':
                fbafee=fbafee+table.row_values(i)[21]
            elif table.row_values(i)[2]=='debt':
                debt=debt+table.row_values(i)[21]
            elif table.row_values(i)[2]=='Transfer':
                print '提现',table.row_values(i)
                writer.writerow([table.row_values(i)])
        sum=income+fee+debt+refund+fbafee
        name=''
        for num in range(len(allDir)):
            if allDir[num]=='（':
                name=name+'s'+'e'
        writer.writerow([allDir,round(income,2),round(fee,2),debt, refund, fbafee, round(sum,2)])
        writer.writerow(['-----------------'])
        print allDir,'营业收入',income,' 营业费用',fee,' 债务',debt,' 退款',refund,' FBA费用',fbafee,' 应收款',sum
        print

    elif child[-3:]=='csv':
        income = 0
        fee = 0
        debt = 0
        refund = 0
        fbafee = 0
        sum = 0
        transfer = 0
        table = open(child,'r')
        lines = table.readlines()
        for i in range(len(lines)):
            if lines[i].split(',')[3]=='Order':
                income=income+float(lines[i].split(',')[16])+float(lines[i].split(',')[17])+float(lines[i].split(',')[18])+float(lines[i].split(',')[19])
                fee=fee+float(lines[i].split(',')[20])+float(lines[i].split(',')[21])+float(lines[i].split(',')[22])+float(lines[i].split(',')[23])
            elif lines[i].split(',')[3]=='Refund' or lines[i].split(',')[3]=='Ajustment' or lines[i].split(',')[3]=='ATOZ':
                refund=refund+float(lines[i].split(',')[25])
            elif lines[i][2].split(',')=='FBA Inventory Fee' or lines[i][2].split(',')=='Service Fee':
                fbafee=float(fbafee+lines[i].split(',')[25])
            elif lines[i][2].split(',')=='debt':
                debt=debt+float(lines[i].split(',')[25])
            elif lines[i].split(',')[3]=='Transfer':
                writer.writerow([lines[i]])
                print '提现',lines[i]
        sum = income + fee + debt + refund + fbafee
        writer.writerow([allDir, round(income, 2), round(fee, 2), debt, refund, fbafee, round(sum, 2)])
        writer.writerow(['-----------------'])
        print allDir,'营业收入', income, ' 营业费用', fee, ' 债务', debt, ' 退款', refund, ' FBA费用', fbafee, ' 应收款', sum
        print