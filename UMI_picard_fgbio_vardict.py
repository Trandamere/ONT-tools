#!/usr/bin/env python3
# coding:utf-8
# conda install -c bioconda vardict
# conda install -c bioconda samtools=1.9
# conda install -c bioconda bwa

import argparse
import os

parser = argparse.ArgumentParser(description="概述:肿瘤UMI")
parser.add_argument('-picard', "--picard", required=True,  help="输入文件:picard.jar所在路径")
parser.add_argument('-fgbio', "--fgbio", required=True,  help="输入文件:fgbio.jar所在路径")
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
cmd1=f'''java -Xmx8G -jar {args.picard} FastqToSam FASTQ={args.FQ1} FASTQ2={args.FQ2} OUTPUT={out_dir}{args.prefix}.uBAM READ_GROUP_NAME={args.prefix} SAMPLE_NAME={args.prefix} LIBRARY_NAME={args.prefix} PLATFORM_UNIT=HiseqX10  PLATFORM=illumina RUN_DATE=`date --iso-8601=seconds`
'''
print(cmd1)
os.system(cmd1)
print(f'''{args.prefix} uBAM Done''')

cmd2=f'''java -Xmx8G -jar {args.fgbio} ExtractUmisFromBam --input={out_dir}{args.prefix}.uBAM --output={out_dir}{args.prefix}.umi.uBAM --read-structure=2M148T 2M148T --single-tag=RX --molecular-index-tags=ZA ZB
'''
print(cmd2) 
os.system(cmd2)
print(f'''{args.prefix} UMI uBAM Done''')

cmd3=f'''samtools fastq {out_dir}{args.prefix}.umi.uBAM | bwa mem -t 8 -p {args.ref} /dev/stdin| samtools view -b > {out_dir}{args.prefix}.umi.BAM
'''
print(cmd3)
os.system(cmd3)
print(f'''{args.prefix} UMI BAM Done''')

cmd4=f'''java -Xmx8G -jar {args.picard} MergeBamAlignment R={args.ref} UNMAPPED_BAM={out_dir}{args.prefix}.umi.uBAM ALIGNED_BAM={out_dir}{args.prefix}.umi.BAM O={out_dir}{args.prefix}.umi.merged.BAM CREATE_INDEX=true MAX_GAPS=-1 ALIGNER_PROPER_PAIR_FLAGS=true VALIDATION_STRINGENCY=SILENT SO=coordinate ATTRIBUTES_TO_RETAIN=XS'''
print(cmd4)
os.system(cmd4)
print(f'''{args.prefix} UMI mergeBAM Done''')

cmd5=f'''java -Xmx8G -jar {args.fgbio} GroupReadsByUmi --input={out_dir}{args.prefix}.umi.merged.BAM --output={out_dir}{args.prefix}.umi.group.BAM --strategy=paired  --min-map-q=20  --edits=1 --raw-tag=RX'''
print(cmd5)
os.system(cmd5)
print(f'''{args.prefix} UMI group Done''')

cmd6=f'''java -Xmx8G -jar {args.fgbio}  CallMolecularConsensusReads --min-reads=1 --min-input-base-quality=20 --input={out_dir}{args.prefix}.umi.group.BAM --output={out_dir}{args.prefix}.consensus.uBAM
'''
print(cmd6)
os.system(cmd6)
print(f'''{args.prefix} call UMI group Done''')

cmd7=f'''samtools fastq {out_dir}{args.prefix}.consensus.uBAM | bwa mem -t 8 -p {args.ref}  /dev/stdin | samtools view -b > {out_dir}{args.prefix}.consensus.BAM'''
print(cmd7)
os.system(cmd7)
print(f'''{args.prefix} map consensus uBAM Done''')

cmd8=f'''java -Xmx8G -jar {args.picard} MergeBamAlignment R={args.ref} UNMAPPED_BAM={out_dir}{args.prefix}.consensus.uBAM ALIGNED_BAM={out_dir}{args.prefix}.consensus.BAM O={out_dir}{args.prefix}.consensus.merge.BAM CREATE_INDEX=true MAX_GAPS=-1 ALIGNER_PROPER_PAIR_FLAGS=true VALIDATION_STRINGENCY=SILENT SO=coordinate ATTRIBUTES_TO_RETAIN=XS
'''
print(cmd8)
os.system(cmd8)
print(f'''{args.prefix} consensus merge BAM Done''')

cmd9=f'''java -Xmx8G -jar {args.fgbio} FilterConsensusReads --input={out_dir}{args.prefix}.consensus.merge.BAM --output={out_dir}{args.prefix}.consensus.merge.filter.BAM --ref={args.ref} --min-reads=2 --max-read-error-rate=0.05 --max-base-error-rate=0.1 --min-base-quality=30 --max-no-call-fraction=0.20 
'''
print(cmd9)
os.system(cmd9)
print(f'''{args.prefix} filter consensus BAM Done''')

cmd10=f'''java -jar {args.fgbio} ClipBam --input={out_dir}{args.prefix}.consensus.merge.filter.BAM --output={out_dir}{args.prefix}.consensus.merge.filter.clip.BAM --ref={args.ref} --clip-overlapping-reads=true
'''
print(cmd10)
os.system(cmd10)
print(f'''{args.prefix} clip consensus BAM Done''')

cmd11=f'''vardict -G {args.ref} -N {args.prefix} -b {out_dir}{args.prefix}.consensus.merge.filter.clip.BAM -z -c 1 -S 2 -E 3 -g 4 -th 4 {args.target} |teststrandbias.R | var2vcf_valid.pl -N test -E -f 0.01 >{out_dir}{args.prefix}.vcf'''
print(cmd11)
os.system(cmd11)
print(f'''{args.prefix} call snp Done''')