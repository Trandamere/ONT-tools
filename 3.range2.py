#!/usr/bin/env python3
# coding:utf-8
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：计算嵌合序列重叠区域数量")
parser.add_argument('-i', "--input_file", required=True,  help="嵌合体的minimap2比对结果")
parser.add_argument('-o', "--output_dir", required=True,  help="计算结果输出文件夹")
parser.add_argument('-n', "--barcode_number", required=True,  help="barcode编号")
parser.add_argument('-O', "--min_covergae", type=float, default=40.0,help="判定为同一嵌合区域的最低覆盖度")
args = parser.parse_args()
def ChimRegionCount(x,y,z,Ox):
    myfile=open('%sLength-Percent-Chimeric.minimap2.result.taxonomy'%y,'w')
    myfile2=open('%sChimeric-seq-Chimeric.minimap2.result.taxonomy'%y,'w')
    myfile3=open('%sChimeric.minimap2.result.taxonomy.gtf'%y,'w')
    Chem=[];ChemSeq=[];LineInfo=[]
    for line1 in open(x):
        ID=line1.split('\t')[0]
        Info=line1.split('\t')
        LineInfo.append(Info)
        if ID not in Chem:
            Chem.append(ID)
        TotalLength=float(line1.split('\t')[1]) 
        start=float(line1.split('\t')[2])/TotalLength*100
        stop =float(line1.split('\t')[3])/TotalLength*100
        SeqID=line1.split('\t')[5]
        Location='%s-%s'%((line1.split('\t')[2],line1.split('\t')[3]))
        IDandLocation='%s_%s'%(ID,Location)
        if IDandLocation not in ChemSeq:
            ChemSeq.append(IDandLocation)
    #首选第一个为原始区域#将下一个区域和原始区域比，重叠度大于40的，判断为同一区域#小于40的判断为其他区域
    #将小于40的其他区域加入原始区域，进行下一个区域判断，到所有结果结束#判断原始区域分成几块#大于1快的就是真的嵌合序列
    myfile.write('Barcode\tSeqName\tstart-stop\tChimeric-length\tFullLength\tPercent\tClassifyNumber\n')
    myfile2.write('SeqID\tSeqkeys\tSeqkeylist\tStart\tStop\tLength\n')
    QunDuan=[]
    SamTax=[];SamRef=[]
    for i in Chem:
        ChemNum=[]
        for j in ChemSeq:
            if '%s_'%i in j:
                ChemNum.append((int(j.split('_')[1].split('-')[0]),int(j.split('_')[1].split('-')[1])))
        ChemNum=sorted(ChemNum)
        if len(ChemNum)==1:#只有一个匹配位置#不能鉴定为嵌合序列
            QunDuan.append(1)
        elif len(ChemNum)>1:#初步匹配到两个位置的可能是嵌合序列，还需要继续进行覆盖度比对来区分
            Classify={}
            #遍历键值进行分类存储
            #接下来涉及起始和终止位点的区域比较
            Classify[ChemNum[0]]=ChemNum[0]
            del ChemNum[0]
            while len(ChemNum)>0:
                for i2 in ChemNum:
                    Cumulate=[]
                    for key,value in Classify.items():
                        Overlap=[]
                        for i3 in range(i2[0],i2[1]+1):
                            if i3 in range(key[0],key[1]+1):
                                Overlap.append(i3)
                        Length=key[1]-key[0]
                        OverlapPercent=float(len(Overlap))/Length*100
                        #直到循环结束才加新的
                        if Ox<=OverlapPercent<=101.00:#如果有匹配的结果，当场添加并跳出循环#40的部分可作为参数开放
                            value2=[value]
                            value2.append((i2))
                            Classify[key]=value2#添加所有已有的元素
                            ChemNum.remove(i2)
                            break
                        else:
                            Cumulate.append('No')#没有现有范围的结果，进行累计，累计到所有结果都没有现有归类，新开一个
                    if len(Cumulate)==len(Classify):#如果遍历字典所有结果依然没有匹配结果添加键值
                        Classify[i2]=i2
                        ChemNum.remove(i2)
                        break
            QunDuan.append(len(Classify))
            #print i,'Classify',Classify,len(Classify)

            if len(Classify)>=2:
                Ref={};Taxon={}#存参考序列#存放参考物种
                for key,value in Classify.items():
                    FullLength=[]
                    #print key#key的长度可以作为代表，计算嵌合部分长度，再除以序列全长就是，嵌合部分长度百分比
                    for IA in LineInfo:
                        if i==IA[0]:
                            FullLength.append(int(IA[1]))
                    #print i,key,value
                    #print 'value',value      
                    #print i,key,key[1]-key[0],FullLength[0],'%s'%(round(float(key[1]-key[0])/FullLength[0]*100,2))+'%'
                    myfile.write('%s\t%s\t%s\t%s\t%s\t%s\t%s'%(z,i,key,key[1]-key[0],FullLength[0],(round(float(key[1]-key[0])/FullLength[0]*100,2)),len(Classify))+'\n')
                    value2=str(value)
                    value3= value2.replace('[','').replace(']','').replace(', (','\t(').split('\t')
                    #print 'value3',value3
                    
                    for IC in value3:#2区域大于2的序列名#序列，物种比对
                        RefInf=[];TaxonInf=[]
                        for IB in LineInfo:
                            if i==IB[0]:
                                #print i,IB
                                #print i,key,value,'IC',IC,IC.split(',')[0].split('(')[1]#提取值中的序列名和物种存入字典，随后进行循环比较，看有没有一样的
                                START=IC.split(',')[0].split('(')[1];STOP=IC.split(',')[1].split(')')[0]
                                #print START,STOP,IB[2],IB[3]
                                #print START==IB[2]
                                #print STOP.strip()==IB[3].strip()
                                
                                if START==str(IB[2]):
                                    if STOP.strip()==str(IB[3]):
                                        #print i,key,value,'IC',IC,IB
                                        RefInf.append(IB[5])#将同一位点的参考序列名字添加入列表中以备后续和范围生成的键存入字典
                                        TaxonInf.append(IB[-1])
                        #print i,IC,RefInf,TaxonInf
                        Ref[IC]=RefInf;Taxon[IC]=TaxonInf#将参考序列和参考物种的信息，通过位置的键值存入字典
                #print i,Ref,Taxon
                #接下来遍历键值进行同一参考序列搜索
                Already=[]#防止前后换位重叠统计
                for key4,value4 in Ref.items():
                    for key5,value5 in Ref.items():
                        if key4!=key5:#将不同区域的键选出进行比较#如果键不一样，遍历值进行不同键相同值得寻找
                            for IE in value4:
                                for IF in value5:
                                    if '%s-%s'%(IE,IF) not in Already:
                                        if IE==IF:
                                            pass#print i,key4,key5,IE,IF
                                            Already.append('%s-%s'%(IE,IF))#防止前后换位重叠统计
                                            Already.append('%s-%s'%(IF,IE))#防止前后换位重叠统计
                                            SamRef.append(i)
                Already2=[]
                for key6,value6 in Taxon.items():
                    for key7,value7 in Taxon.items():
                        if key6!=key7:
                            for IG in value6:
                                for IH in value7:
                                    if '%s-%s'%(IG,IH) not in Already2:
                                        if IG==IH:
                                            #print i,key6,key7,IG,IH
                                            Already2.append('%s-%s'%(IG,IH))#防止前后换位重叠统计
                                            Already2.append('%s-%s'%(IH,IG))#防止前后换位重叠统计
                                            SamTax.append(i)
                #此处进行嵌合部分序列区域定位
                #嵌合区域定位，将key的健收集，排序，后放入列表中，删除奇数元素，剩余元素，每一对为嵌合区域
                ChimKey=[];Strand=[]
                for IM in LineInfo:
                    if i==IM[0]:
                        Strand.append(IM[4])#print i,IM
                for key8,value8 in Classify.items():
                    ChimKey.append(key8)
                ChimKey.sort()
                ChimKeylist=[]
                for IJ in ChimKey:
                    for IK in IJ:
                        ChimKeylist.append(IK)
                N=1
                while N<=len(ChimKeylist)-2:#2,3#5,6
                    if ChimKeylist[N+1]-ChimKeylist[N]>=20:
                        #print i,ChimKey,ChimKeylist,ChimKeylist[N],ChimKeylist[N+1],ChimKeylist[N+1]-ChimKeylist[N]
                        #print '%s\t%s\t%s\t%s\t%s\t%s\n'%(i,ChimKey,ChimKeylist,ChimKeylist[N],ChimKeylist[N+1],ChimKeylist[N+1]-ChimKeylist[N])
                        myfile2.write('%s\t%s\t%s\t%s\t%s\t%s\n'%(i,ChimKey,ChimKeylist,ChimKeylist[N],ChimKeylist[N+1],ChimKeylist[N+1]-ChimKeylist[N]))
                        #print '%s\tSTGJ\texon\t%s\t%s\t.\t%s\t.\ttranscript_id "%s"\n'%(i,ChimKeylist[N],ChimKeylist[N+1],Strand[0],i)
                        myfile3.write('%s\tSTGJ\texon\t%s\t%s\t.\t%s\t.\ttranscript_id "%s";\n'%(i,ChimKeylist[N],ChimKeylist[N+1],Strand[0],i))
                    N+=2
                    
                #3嵌合部分提取
            #遍历字典的键，用键搜索序列名称及长度，计算嵌合序列各部分占长度百分比
            #首先判断字典值长

    myfile.close()
    myfile2.close()
    myfile3.close()
ChimRegionCount(args.input_file,args.output_dir,args.barcode_number,args.min_covergae)
#ChimRegionCount('Chimeric-BC01_M01.minimap2.result.taxonomy')

