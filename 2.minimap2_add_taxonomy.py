#!/usr/bin/env python3
# coding:utf-8
import re
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：对minimap2的比对结果增加taxonomy lineage信息")
parser.add_argument('-i', "--map_result", required=True,  help="输入文件：minimap2比对结果文件")
parser.add_argument('-a', "--acc_id2taxid", required=True,  help="输入文件：acc_id2taxid文件")
parser.add_argument('-n', "--taxid2name", required=True, help="输入文件：taxid2name文件")
parser.add_argument('-o', "--out", required=True,  help="输出文件名称，可以包含路径名称，如果路径不存在会自动创建")
args = parser.parse_args()

out_dir = os.path.dirname(args.out)
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)

dict_2 = {}
with open (args.acc_id2taxid,'r',encoding='utf-8') as f:
    for line in f.readlines():
        acc_id,tax_id = line.strip().split("\t")
        dict_2[acc_id] = tax_id
dict_3 = {}
with open (args.taxid2name,'r',encoding='utf-8') as f:
    for line in f.readlines():
        tax_id,name_info = line.strip().split("\t")
        dict_3[tax_id] = name_info
output = open (args.out,'w')
with open (args.map_result,'r',encoding='utf-8') as file:
    for line in file.readlines():
        line = line.strip()
        if line:
            #de:f为sequence divergence = 1 - Gap-compressed identity
            if re.search("(de:f:\d+.*?)\t",line):
                de_f = re.search("(de:f:\d+.*?)\t",line).group(1)
                line_list = line.split("\t")
                acc_id = line_list[5]
                if acc_id in dict_2:
                    tax_id = dict_2[acc_id]
                    if tax_id in dict_3:
                        name_info = dict_3[tax_id]
                        info = "\t".join(line_list[0:11]) + "\t" + line_list[14] + "\t" + de_f
                        output.write(info + "\t" + tax_id + "\t" + name_info + "\n")
                    else:
                        print("tax_id not in database:" + tax_id)
                else:
                    print("acc_id not in database:" + acc_id)
            else:
                print(line)
output.close()
