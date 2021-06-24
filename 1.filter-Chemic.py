# -*- coding:utf-8 -*-
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：寻找嵌合体对应的minimap结果")
parser.add_argument('-i1', "--input_dir1", required=True,  help="yacrd出处原始结果")
parser.add_argument('-i2', "--input_dir2", required=True,  help="minimap2比对结果")
parser.add_argument('-i3', "--input_dir3", required=True,  help="输出文件夹位置")
args = parser.parse_args()
def Chem(x,y,z):
    myfile=open('%s/Chimeric-minimap2.result'%z,'w')
    ID=[]
    Chrem='Chimeric\t'
    for line1 in open(x):
        if Chrem in line1:
            ID.append(line1.split('\t')[1])
            #print(line1.split('\t')[1])
    for i in ID:
        for line2 in open(y):
            if '%s\t'%i in line2:
                IDmap=(line2.split('\t')[0:6])
                #myfile.write('Chimeric\t%s'%line2)
                myfile.write(line2)
                #print '\t'.join(IDmap)
    myfile.close()
Chem(args.input_dir1,args.input_dir2,args.input_dir3)
#Chem('barcode01-report.yacrd','BC01_M01.minimap2.result')
#Chem('barcode02-report.yacrd','BC02_M02.minimap2.result')
#Chem('barcode03-report.yacrd','BC03_M03.minimap2.result')
#Chem('barcode04-report.yacrd','BC04_M04.minimap2.result')
#Chem('barcode05-report.yacrd','BC05_M05.minimap2.result')
#Chem('barcode06-report.yacrd','BC06_M06.minimap2.result')
