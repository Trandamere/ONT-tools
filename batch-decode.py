#!/usr/bin/env python3
# coding:utf-8
#example python batch-decode.py -i /mnt/d/encode/guangzhoulab001-20211223-test1/
import argparse
import os
import sys

parser = argparse.ArgumentParser(description="功能:寻找目录下所有压缩文件，输入密码进行解压")
parser.add_argument('-i', "--input_dir", required=True,  help="barcode所在目录")
parser.add_argument('-o', "--out", default='NA',  help="输出结果路径，如果路径不存在会自动创建")
parser.add_argument('-p', "--password", default='Shengting@123', help="解压密码")
args = parser.parse_args()
code_dir = sys.path[0] + "/"
# out_dir = os.path.dirname(args.out)
# if out_dir and not os.path.exists(out_dir):
#     os.makedirs(out_dir)
out_dir =args.out
print('out_dir',out_dir)
def Decode(input,output,password):
    for root,dirs,files in os.walk(input):
        #print(root,dirs,files)
        for Files in (files):
            if 'fastq.gz' in Files:
                print('root',root)
                print('Files',Files)
                print('input',input)
                Spath =root.replace(input,'')
                print('Spath',Spath)
                BarcodeString=root.split('/')
                BarNumL=[]
                for i in BarcodeString:
                    if 'barcode' or 'unclassified' in i:
                        BarNumL.append(i)
                BarcodeNum=BarNumL[0]
                
                outpath=output+'/'+BarcodeNum+'/'
                outfile=output+'/'+BarcodeNum+'/'+Files
                if out_dir=='NA':
                    try:
                        cmd = '7z x -p'+password+' '+input+Spath+BarcodeNum+'/'+Files +' -o'+input+Spath+BarcodeNum
                    except:
                        cmd = ''
                    # print('cmd',cmd)
                    os.system(cmd)
                    try:#删除加密的gz
                        cmd2 = 'rm '+input+Spath+BarcodeNum+'/'+Files
                    except:
                        cmd2 = ''
                    print('cmd2',cmd2)
                    os.system(cmd2)
                    #进行压缩
                    try:
                        cmd3='7z a -tgzip %s.gz %s'%(input+Spath+BarcodeNum+'/'+Files.split('.gz')[0],input+Spath+BarcodeNum+'/'+Files.split('.gz')[0])
                    except:
                        cmd3=''
                    # print('cmd3',cmd3)
                    os.system(cmd3)
                    try:#s删除q
                        cmd4 = 'rm '+input+Spath+BarcodeNum+'/'+Files.split('.gz')[0]
                    except:
                        cmd4 =''
                    # os.system(cmd4)
                else:#输出了目录
                    Spath =root.replace(input,'')
                    out_dir2 = os.path.dirname(args.out)
                    if out_dir2 and not os.path.exists(out_dir2):
                        os.makedirs(out_dir2)
                    try:
                        cmd = '7z x -p'+password+' '+input+Spath+BarcodeNum+'/'+Files +' -o'+out_dir2+'/'+Spath+BarcodeNum
                    except:
                        cmd = ''
                    print('cmd',cmd)
                    os.system(cmd)
                    # try:#删除加密的gz
                    #     cmd2 = 'rm '+input+Spath+BarcodeNum+'/'+Files
                    # except:
                    #     cmd2 = ''
                    # print('cmd2',cmd2)
                    # os.system(cmd2)
                    #进行压缩
                    try:
                        cmd3='7z a -tgzip %s.gz %s'%(out_dir2+'/'+Spath+BarcodeNum+'/'+Files.split('.gz')[0],out_dir2+'/'+Spath+BarcodeNum+'/'+Files.split('.gz')[0])
                    except:
                        cmd3=''
                    # print('cmd3',cmd3)
                    os.system(cmd3)
                    try:#s删除q
                        cmd4 = 'rm '+out_dir2+'/'+Spath+BarcodeNum+'/'+Files.split('.gz')[0]
                    except:
                        cmd4 =''
                    os.system(cmd4)
    print('解密完成')
                
Decode(args.input_dir,out_dir,args.password)
