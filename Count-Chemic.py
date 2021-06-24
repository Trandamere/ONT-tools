# -*- coding:utf-8 -*-
import argparse
import os
import sys
parser = argparse.ArgumentParser(description="功能：进行嵌合序列部分信息统计")
parser.add_argument('-i', "--input_dir", required=True,  help="嵌合序列结果所在barcode在目录")
parser.add_argument('-o', "--out", required=True,  help="统计结果输出路径，可以包含路径名称，如果路径不存在会自动创建")
args = parser.parse_args()
out_dir = os.path.dirname(args.out)
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)
def Count(x,y):
    Fold=[]
    for root,dirs,files in os.walk(r"%s"%x):
        for file in files:
            Fold.append(root)
    Fold=sorted(list(set(Fold)))
    myfile=open(y+'/total-Chimeric-range-count.txt','w')
    myfile2=open(y+'/total-Chimeric-same-Ref.txt','w')
    myfile3=open(y+'/total-Chimeric-same-taxiom.txt','w')
    myfile.write('barcode\tRange2\tRange>2\n')
    for i in Fold:#

        if 'barcode' in i:
                #print(i) 
                BarcodeNum=i.split('/')[-1]
                BarcodeFold='%s/%s'%(out_dir,BarcodeNum)
                #print(BarcodeFold)
                if BarcodeFold and not os.path.exists(BarcodeFold):
                    os.makedirs(BarcodeFold)
                BarcodeNum2=BarcodeNum.split('barcode')[1]
                #print(i,BarcodeNum,BarcodeFold,BarcodeNum2)
                #print('seqkit fx2tab '+i+'/'+BarcodeNum+'.fastq '+'-o '+y+BarcodeNum+'.fastq.count.txt')
                '''
                os.system('seqkit fx2tab '+i+'/'+BarcodeNum+'.fastq '+'-o '+y+'/'+BarcodeNum+'/'+BarcodeNum+'.fastq.count.txt')
                os.system('wc -l '+y+BarcodeNum+'/'+BarcodeNum+'.fastq.count.txt >>'+y+'/total-seq.txt')
                os.system('grep "Chimeric" '+i+'/'+BarcodeNum+'-report.yacrd > '+y+'/'+BarcodeNum+'/'+BarcodeNum+'count-Chimeric.txt')
                os.system('wc -l '+y+'/'+BarcodeNum+'/'+BarcodeNum+'count-Chimeric.txt >>'+y+'/total-Chimeric.txt')
                os.system("cat "+i+'/'+"Chimeric-minimap2.result |awk '{print $1}'| sort | uniq >"+y+BarcodeNum+'/'+BarcodeNum+'Chimeric-minimap2-count.txt')
                print("cat "+i+'/'+"Chimeric-minimap2.result |awk '{print $1}'| sort | uniq >"+y+BarcodeNum+'/'+BarcodeNum+'Chimeric-minimap2-count.txt')
                os.system('wc -l '+y+BarcodeNum+'/'+BarcodeNum+'Chimeric-minimap2-count.txt >>'+y+'/total-Chimeric-have-minimap2.txt')
                '''
                ##统计嵌合区域数量
                '''
                Chim2=[];Chim3=[]
                for line1 in open (i+'/Length-Percent-Chimeric.minimap2.result.taxonomy'):
                    if 'SeqName' not in line1:
                        if int(line1.strip().split('\t')[-1])==2:
                            Chim2.append(line1.split('\t')[1])
                        if int(line1.strip().split('\t')[-1])>2:
                            Chim3.append(line1.split('\t')[1])
                Chim2=list(set(Chim2));Chim3=list(set(Chim3))
                #print(BarcodeNum,'Chim2',len(Chim2),'Chim3',len(Chim3))
                myfile.write(BarcodeNum+'\t'+str(len(Chim2))+'\t'+str(len(Chim3))+'\n')
                '''
    
                ##统计是否来源同一参考序列及同一物种
                #print(i+'/'+f'Chimeric-BC{BarcodeNum2}_M{BarcodeNum2}.minimap2.result.taxonomy')
                Chem=[];ChemSeq=[];LineInfo=[]
                for line2 in open(i+'/'+f'Chimeric-BC{BarcodeNum2}_M{BarcodeNum2}.minimap2.result.taxonomy'):
                    ID=line2.split('\t')[0]
                    Info=line2.split('\t')
                    LineInfo.append(Info)
                    if ID not in Chem:
                        Chem.append(ID)
                    TotalLength=float(line2.split('\t')[1]) 
                    start=float(line2.split('\t')[2])/TotalLength*100
                    stop =float(line2.split('\t')[3])/TotalLength*100
                    SeqID=line2.split('\t')[5]
                    Location='%s-%s'%((line2.split('\t')[2],line2.split('\t')[3]))
                    IDandLocation='%s_%s'%(ID,Location)
                    if IDandLocation not in ChemSeq:
                        ChemSeq.append(IDandLocation)
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
                                    if 40.00<=OverlapPercent<=101.00:#如果有匹配的结果，当场添加并跳出循环#40的部分可作为参数开放
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
                                myfile.write('%s\t%s\t%s\t%s\t%s\t%s'%(x.split('-')[1].split('_')[0],i,key,key[1]-key[0],FullLength[0],(round(float(key[1]-key[0])/FullLength[0]*100,2)))+'%\n')
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
                SamTax=list(set(SamTax));SamRef=list(set(SamRef))
                #print('Same Ref',len(SamRef))
                myfile2.write(BarcodeNum+'\t'+str(len(SamRef))+'\n')
                #print('Same Tax',len(SamTax))
                myfile3.write(BarcodeNum+'\t'+str(len(SamTax))+'\n')
    myfile.close()
    myfile2.close()
    myfile3.close()
Count(args.input_dir,args.out)
