import time,subprocess,os,logging
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
import os.path
from pathlib import Path
import base64
F1='.fast5'
F2='.fastq.gz'
# 设定监控日志输出文件名和内容形式
logging.basicConfig(format='%(asctime)s - %(message)s', filename='/opt/ont/SeeGeneOSS/code.log', filemode='a', level=logging.INFO)
if __name__ == "__main__":
    patterns = "*"
    ignore_patterns = ""
    ignore_directories = False
    case_sensitive = True
    my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)
Already=[]#加密fastqgz文件
Nocodefile=[]#所有文件fastqgz
MBpath = []#存放md路径
PassWord = (base64.b64decode("U2hlbmd0aW5nQDEyMw==").decode("utf-8"))
BASEDIR = os.path.abspath(os.path.dirname(__file__))
def get_filename(filename):
    "Get the name of file that created"
    return filename.split('/')[-1]
def get_filepath(filename):
    "Get the name of file that created"
    filepath='/'.join(filename.split('/')[0:-2])
    return filepath

def run_tests(name):
    os.chdir(BASEDIR)
    

def on_created(event):
    if event.src_path.endswith(".fast5"):
        print(f"文件{event.src_path} 被创建")
        # print(str(event.src_path).split('/')[1])
    if event.src_path.endswith(".fastq.gz"):
        print(f"文件{event.src_path} 被创建")
        # print(str(event.src_path).split('/')[1])
 

def on_deleted(event):
    if event.src_path.endswith(".fast5"):
        print(f"文件 {event.src_path} 被删除了")
    if event.src_path.endswith(".fastq.gz"):
        print(f"文件 {event.src_path} 被删除了")

def on_modified(event):
    if event.src_path.endswith(".fast5"):
        print(f"文件 {event.src_path} 被修改了")
    if event.src_path.endswith(".fastq.gz"):
        print(f"文件 {event.src_path} 被删除了")
    
def on_moved(event):
    print(f"文件从 {event.src_path} 移动到 {event.dest_path}")
    # filename=event.dest_path.split('/')[-1]
    if event.dest_path.endswith(".fast5"):
        created_file_size = os.path.getsize(os.path.join(event.dest_path))
        #time.sleep(10)
        new_file_size = os.path.getsize(os.path.join(event.dest_path))
        if created_file_size == new_file_size: 
            # logging.info("移动文件: % s" % os.path.join(event.dest_path))
            try:
                cmd='7z a -p%s %s.7z %s'%(PassWord,os.path.join(event.dest_path),os.path.join(event.dest_path))
                logging.info("% s完成加密" % os.path.join(event.dest_path))
            except:
                cmd=''
                logging.info("% s加密失败" % os.path.join(event.dest_path))
            print(cmd)
            os.system(cmd)
            try:#将文件7z改为gz
                os.system('mv %s.7z %s.gz'%(os.path.join(event.dest_path),os.path.join(event.dest_path)))
            except:
                pass
            try:#移除压缩文件
                os.system('rm %s'%os.path.join(event.dest_path))
            except:
                pass
    if event.dest_path.endswith(".fastq.gz"):
        if event.src_path != event.dest_path:
            if event.dest_path not in Already:
                created_file_size = os.path.getsize(os.path.join(event.dest_path))
                # time.sleep(10)
                new_file_size = os.path.getsize(os.path.join(event.dest_path))
                if created_file_size == new_file_size: 
                    # logging.info("移动文件: % s" % os.path.join(event.dest_path))
                    filename=get_filename(event.dest_path)
                    print('filename',filename)
                    filepath=event.dest_path.replace(filename,'')
                    print('filepath',filepath)
                    try:
                        cmd='7z x %s -o%s'%(os.path.join(event.dest_path),filepath)
                    except:
                        cmd=''
                    os.system(cmd)
                    print('unzip',cmd)
                    try:
                        os.system('rm %s'%event.dest_path)
                    except:
                        pass
                    try:
                        cmd='7z a -p%s %s.7z %s'%(PassWord,os.path.join(event.dest_path.split('.gz')[0]),os.path.join(event.dest_path.split('.gz')[0]))
                        logging.info("% s完成加密" % os.path.join(event.dest_path))
                        Already.append(event.dest_path)
                    except:
                        cmd=''
                        logging.info("% s加密失败" % os.path.join(event.dest_path))
                    print(cmd)
                    os.system(cmd)
                    try:#将文件7z改为gz
                        os.system('mv %s.7z %s.gz'%(os.path.join(event.dest_path.split('.gz')[0]),os.path.join(event.dest_path.split('.gz')[0])))
                    except:
                        pass
                    try:#移除压缩文件
                        os.system('rm %s'%os.path.join(event.dest_path.split('.gz')[0]))
                    except:
                        pass
                    ####在这里检测summary文件是否存在，进行文件数量核对
                    # if len(Nocodefile)==len(Already):#如果summ里的fastq.gz文件与已经加密的文件数量一样
                    #     Summary_path=event.dest_path.split('fastq_')[0]#进行mb文件生成#此处需要获得summary路径
                    #     cmd5 = 'touch %sFinish.mb'%Summary_path
                    #     logging.info(cmd5)
                    #     os.system(cmd5)

                        
    if "sequencing_summary" in event.dest_path and event.dest_path.endswith(".txt"):
        logging.info(f"文件从 {event.src_path} 移动到 {event.dest_path}")
        # time.sleep(1)
        Summary_path=event.dest_path.split('fastq_')[0]
        MBpath.append(Summary_path)
        for line1 in open(event.dest_path):
            if 'fastq.gz' in line1:
                Pass_fastq=line1.split('\t')[0]
                if Pass_fastq not in Nocodefile:
                	Nocodefile.append(Pass_fastq)
        MBname = MBpath[-1].replace(MBpath[-1].split('/')[-1],'')
        time.sleep(5400)
        cmd6 = 'touch %sFinish.mb'%MBname
        #logging.info(cmd6)
        my_file=Path('%sFinish.mb'%MBname)
        if my_file.is_file():
            	pass
        else:
            os.system(cmd6)
            logging.info(cmd6)


my_event_handler.on_created = on_created
my_event_handler.on_deleted = on_deleted
my_event_handler.on_modified = on_modified
my_event_handler.on_moved = on_moved


path = "/var/lib/minknow/data/"
go_recursively = True
my_observer = Observer()
my_observer.schedule(my_event_handler, path, recursive=go_recursively)


my_observer.start()
logging.info(f"start morinor {path}")
Already=list(set(Already))
Nocodefile=list(set(Nocodefile))
while True:
    if len(Already)!=0:
        if len(Already)!=len(Nocodefile):
            print('len no equal',Already,Nocodefile)
            time.sleep(60)
        if len(Already)==len(Nocodefile):
            MBname = MBpath[-1].replace(MBpath[-1].split('/')[-1],'')
            cmd5 = 'touch %sFinish.mb'%MBname
            #logging.info(cmd5)
            my_file=Path('%sFinish.mb'%MBname)
            if my_file.is_file():
            	pass
            else:
            	os.system(cmd5)
            	logging.info(cmd5)
            print('len equal',Already,Nocodefile)
            time.sleep(60)
            continue
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    my_observer.stop()
    my_observer.join()
