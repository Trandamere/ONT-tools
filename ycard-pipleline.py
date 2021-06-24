#!/usr/bin/env python3
# coding:utf-8
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：寻找文件夹下所有barecode文件夹，解压合并barcode中fastq，生成yacrd运行结果")
parser.add_argument('-i', "--input_dir", required=True,  help="barcode所在目录")
parser.add_argument('-o', "--out", required=True,  help="合并且压缩的fastq名称，以.gz结尾，可以包含路径名称，如果路径不存在会自动创建")
parser.add_argument('-a', "--yacrd_threads",   default=4,help="线程数")
parser.add_argument('-c', "--yacrd_conda_env", default="yacrd",help="yacrd conda环境")
parser.add_argument('-m', "--minimap2_result", required=True,  help="minimap2 比对结果所在文件夹")
parser.add_argument('-tid', "--acc_id2taxid", type=str, default="/home/MicroAnalysisDev/workspace/test/lp-test/minimap2.taxid",  help="minimap2 需要的seqID与taxid对应关系，默认：/home/MicroAnalysisDev/database/NanoTNGS_database/minimap2_database/v2.1/minimap2.taxid")
parser.add_argument('-lin', "--taxid2lineage_txt", type=str, default="/home/MicroAnalysisDev/workspace/test/lp-test/minimap2_lineage.txt",  help="taxid对应的lineage关系，默认：/home/MicroAnalysisDev/database/NanoTNGS_database/minimap2_database/v2.1/minimap2_lineage.txt")
parser.add_argument('-O', "--min_covergae", type=float,default=40.0,help="判定为同一嵌合区域的最低覆盖度,默认40.0，即40%")
args = parser.parse_args()
code_dir = sys.path[0] + "/"
Chempy=code_dir + "1.filter-Chemic.py"
minimap2_add_taxonomy_py = code_dir + "2.minimap2_add_taxonomy.py"
ChimRegionCount_py = code_dir + "3.range2.py"
min_covergae=args.min_covergae
acc_id2taxid = args.acc_id2taxid
taxid2lineage_txt = args.taxid2lineage_txt
out_dir = os.path.dirname(args.out)
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)

def yacrd(x,yacrd_threads,sampledir,yacrd_conda_env,minimap2_result):
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
            BarcodeNum2=BarcodeNum.split('barcode')[1]
            cmd=f'''
#!/bin/bash
#PBS -l nodes=1:ppn={yacrd_threads}
#PBS -q batch
#PBS -N {sampledir}.minimap2
#PBS -e {sampledir}.minimap2.errorlog
#PBS -e {sampledir}.minimap2.outlog
#PBS -V
nprocs=`wc -l < $PBS_NODEFILE`
cd $PBS_O_WORKDIR
#source activate {yacrd_conda_env}
minimap2 -x ava-ont -g 500 {BarcodeFold}/{BarcodeNum}.fastq {BarcodeFold}/{BarcodeNum}.fastq > {BarcodeFold}/{BarcodeNum}-overlap.paf
yacrd -i {BarcodeFold}/{BarcodeNum}-overlap.paf -o {BarcodeFold}/{BarcodeNum}-report.yacrd -c 4 -n 0.4 scrubb -i {BarcodeFold}/{BarcodeNum}.fastq -o {BarcodeFold}/{BarcodeNum}.scrubb.fasta
python {Chempy} -i1 {BarcodeFold}/{BarcodeNum}-report.yacrd -i2 {minimap2_result}BC{BarcodeNum2}_M{BarcodeNum2}*minimap2.result -i3 {BarcodeFold}
python {minimap2_add_taxonomy_py} -i {BarcodeFold}/Chimeric-minimap2.result -a {acc_id2taxid} -n {taxid2lineage_txt} -o {BarcodeFold}/Chimeric-BC{BarcodeNum2}_M{BarcodeNum2}.minimap2.result.taxonomy
python {ChimRegionCount_py} -i {BarcodeFold}/Chimeric-BC{BarcodeNum2}_M{BarcodeNum2}.minimap2.result.taxonomy -o {BarcodeFold}/ -n {BarcodeNum} -O {min_covergae}
'''
            fqpbs_name=BarcodeFold+'/'+BarcodeNum+".yacrd.pbs"
            print(fqpbs_name)
            fqpbs = open (fqpbs_name,'w')
            fqpbs.write(cmd)
            fqpbs.close()
            fqpbs_return = os.popen("qsub " + fqpbs_name)
            fqpbs_id = fqpbs_return.read().split('\n')[0]
#随后需要使用鉴定出的结果对minimap2结果进行过滤
yacrd(args.input_dir,args.yacrd_threads,args.input_dir,args.yacrd_conda_env,args.minimap2_result)
