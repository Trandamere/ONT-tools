# -*- coding:utf-8 -*-
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：寻找文件夹下所有barecode文件夹，解压合并barcode中fastq，生成yacrd运行结果")
parser.add_argument('-i', "--input_dir", required=True,  help="barcode所在目录")
parser.add_argument('-o', "--out", required=True,  help="合并且压缩的fastq名称，以.gz结尾，可以包含路径名称，如果路径不存在会自动创建")
args = parser.parse_args()

out_dir = os.path.dirname(args.out)
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)

def yacrd(x):
    Fold=[]
    for root,dirs,files in os.walk(r"%s"%x):
        for file in files:
            Fold.append(root)
    Fold=sorted(list(set(Fold)))
    for i in Fold:#进行文件夹建立，并进行进行目录下解压，合并，生成yacrd结果
        if 'barcode' in i:
            #print(i) 
            BarcodeNum=i.split('/')[-1]
            BarcodeFold='%s/%s'%(out_dir,BarcodeNum)
            #print(BarcodeFold)
            if BarcodeFold and not os.path.exists(BarcodeFold):
                os.makedirs(BarcodeFold)
                #建立文件夹后进行文件夹下文件解压，合并
            os.system("cat %s/*.fastq.gz >%s/%s.fastq.gz"%(i,BarcodeFold,BarcodeNum))
            os.system("gunzip %s/%s.fastq.gz"%(BarcodeFold,BarcodeNum))
            print("minimap2 -x ava-ont -g 500 %s/%s.fastq %s/%s.fastq > %s/%s-overlap.paf"%(BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum))
            os.system("minimap2 -x ava-ont -g 500 %s/%s.fastq %s/%s.fastq > %s/%s-overlap.paf"%(BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum))
            print("yacrd -i %s/%s-overlap.paf -o %s/%s-report.yacrd -c 4 -n 0.4 scrubb -i %s/%s.fastq -o %s/%s.scrubb.fasta"%(BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum))
            os.system("yacrd -i %s/%s-overlap.paf -o %s/%s-report.yacrd -c 4 -n 0.4 scrubb -i %s/%s.fastq -o %s/%s.scrubb.fasta"%(BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum,BarcodeFold,BarcodeNum))
            
yacrd(args.input_dir)
