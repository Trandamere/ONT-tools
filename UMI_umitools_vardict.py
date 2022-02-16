#!/usr/bin/env python3
# coding:utf-8
# 2022.2.15
#python3.7
#conda install -c bioconda -c conda-forge umi_tools
#conda install -c bioconda bwa
#conda install -c bioconda samtools=1.9
#conda install -c bioconda vardict

import argparse
import os

parser = argparse.ArgumentParser(description="概述:肿瘤UMI")
parser.add_argument('-A', "--FQ1", required=True,help="FQ1文件所在路径及文件名")
parser.add_argument('-B', "--FQ2", required=True,help="FQ2文件所在路径及文件名")
parser.add_argument('-o', "--outdir", required=True,help="输出文件目录")
parser.add_argument('-p', "--prefix", required=True,help="输出文件前缀")
parser.add_argument('-r', "--ref", required=True,help="参考基因组fasta")
parser.add_argument('-t', "--target", required=True,help="目标bed")
args = parser.parse_args()
out_dir=args.outdir
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)

cmd1=f'''umi_tools extract -I {args.FQ1} --bc-pattern=NNNXXXXNN --read2-in={args.FQ2} --stdout {out_dir}'extracted'_{args.FQ1.split('/')[-1]} --read2-out={out_dir}'extracted'_{args.FQ2.split('/')[-1]} --ignore-read-pair-suffixes
'''
print(cmd1)
os.system(cmd1)
print(f'''{args.prefix} extract UMI Done''')

cmd2=f'''bwa mem -t 8 {args.ref} {out_dir}'extracted'_{args.FQ1.split('/')[-1]} {out_dir}'extracted'_{args.FQ2.split('/')[-1]} >{out_dir}{args.prefix}.sam'''
print(cmd2)
os.system(cmd2)
print(f'''{args.prefix} mapping to genome''')

cmd3 = f'''samtools import {args.ref} {out_dir}{args.prefix}.sam {out_dir}{args.prefix}.bam'''
print(cmd3)
os.system(cmd3)
print(f'''{args.prefix} sam to bam''')

cmd4 = f'''samtools sort {out_dir}{args.prefix}.bam -o {out_dir}sort_{args.prefix}.bam'''
os.system(cmd4)

cmd5 = f'''samtools index {out_dir}sort_{args.prefix}.bam'''
os.system(cmd5)

cmd6 = f'''umi_tools dedup -I {out_dir}sort_{args.prefix}.bam --output-stats=deduplicated -S {out_dir}{args.prefix}.deduplicated.bam'''
print(cmd6)
os.system(cmd6)
print(f'''{args.prefix} bam dedup''')

cmd7 = f'''samtools index {out_dir}{args.prefix}.deduplicated.bam'''
os.system(cmd7)

cmd8 = f'''vardict -G {args.ref} -N {args.prefix} -b {out_dir}{args.prefix}.deduplicated.bam -z -c 1 -S 2 -E 3 -g 4 -th 4 {args.target} |teststrandbias.R | var2vcf_valid.pl -N test -E -f 0.01 >{out_dir}{args.prefix}.vcf'''
print(cmd8)
os.system(cmd8)
print(f'''{args.prefix} call snp Done''')