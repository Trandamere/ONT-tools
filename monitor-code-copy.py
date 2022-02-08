#conda install watchdog
#conda install -c conda-forge yagmail
#conda install -c conda-forge premailer
# sudo vim /etc/sysctl.conf
#fs.inotify.max_user_watches=99999999
import time,subprocess,os,logging
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
import os.path
from pathlib import Path
import base64
import yagmail

# 设定监控日志输出文件名和内容形式
logging.basicConfig(format='%(asctime)s - %(message)s', filename='/opt/ont/SeeGeneOSS/code.log', filemode='a', level=logging.INFO)
#logging.basicConfig(format='%(asctime)s - %(message)s', filename='/mnt/c/Users/luping/Desktop/code.log', filemode='a', level=logging.INFO)
if __name__ == "__main__":
    patterns = "*"
    ignore_patterns = ""
    ignore_directories = False
    case_sensitive = True
    my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)
    my_event_handler2 = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)
Already=[]#加密fastqgz文件
Nocodefile=[]#所有文件fastqgz
MBpath = []#存放md路径
PassWord = (base64.b64decode("U2hlbmd0aW5nQDEyMw==").decode("utf-8"))

#链接邮箱服务器
# emails='ping_lu@seegeno.com,changxiao_xie@shengtinggroup.com,qiankyun_li@seegeno.com,xiao_xiong@seegeno.com'
emails='ping_lu@seegeno.com'
yag = yagmail.SMTP(user="noreply@shengtinggroup.com", password="Aa1q2w3e", host='smtphz.qiye.163.com')
# result_set_list = emails.split(',')
result_set_list = emails
set_name =''
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
    

def on_created(event):#创建文件文件路径不在测序数据产生文件夹
    
    if '/var/lib/minknow/data' not in  event.src_path:#产生文件路径名称为fast5或fastq.gz
        if event.src_path.endswith(".fast5") or event.src_path.endswith(".fastq.gz"):###fastq.gz在/var/lib/minknow/之外被创建，就是文件被拷贝到其他地方
            print(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作")
            print('create')
            logging.info(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作")
            # logging.info(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作")#此处添加邮件警告内容
            contents = f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作"
            yag.send(result_set_list, '文件异常复制警示！', contents)
            time.sleep(10)

 

def on_deleted(event):#删除文件
    if event.src_path.endswith(".fast5") or event.src_path.endswith(".fastq.gz"):
        print(f"文件 {event.src_path} 被删除了")
        logging.info(f"文件 {event.src_path} 被删除了")#此处添加邮件警告内容


def on_modified(event):
    
    if '/var/lib/minknow/data' not in event.src_path:
        if event.src_path.endswith(".fast5") or event.src_path.endswith(".fastq.gz"):###fastq.gz在/var/lib/minknow/之外被创建，就是文件被拷贝到其他地方
            print(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作")
            print('modified')
            print(event.src_path)
            logging.info(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作")#此处添加邮件警告内容
            
            contents = f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作"
            #try:
            #    yag.send(result_set_list, '文件异常复制警示！', contents)
            #except:
            #    logging.info(f"警告！ 测序数据 被拷贝至{event.src_path}，请注意测序数据是否为工作人员操作,发送邮件失败")#此处添加邮件警告内容
            time.sleep(10)
    
def on_moved(event):
    if event.src_path.endswith(".fast5") or event.src_path.endswith(".fastq.gz"):
        print(f"文件 {event.src_path} 被移动了")
        logging.info(f"文件 {event.src_path} 被移动了")#此处添加邮件警告内容


# my_event_handler.on_created = on_created
# my_event_handler.on_deleted = on_deleted
# my_event_handler.on_modified = on_modified
# my_event_handler.on_moved = on_moved
my_event_handler2.on_created = on_created
my_event_handler2.on_deleted = on_deleted
my_event_handler2.on_modified = on_modified
my_event_handler2.on_deleted = on_deleted
# path = "/var/lib/minknow/data/"
go_recursively = True
# my_observer = Observer()
# my_observer.schedule(my_event_handler, path, recursive=go_recursively)

path2 = "/media/guangzhoulab-001/"
#path2 ='/mnt/c/Users/luping/'
my_observer2 = Observer()
my_observer2.schedule(my_event_handler2, path2, recursive=go_recursively)

# my_observer.start()
# logging.info(f"start morinor {path}")
Already=list(set(Already))
Nocodefile=list(set(Nocodefile))

my_observer2.start()
logging.info(f"start morinor {path2} for copy")

# while True:
#     if len(Already)!=0:
#         if len(Already)!=len(Nocodefile):
#             print('len no equal',Already,Nocodefile)
#             time.sleep(60)
#         if len(Already)==len(Nocodefile):
#             MBname = MBpath[-1].replace(MBpath[-1].split('/')[-1],'')
#             cmd5 = 'touch %sFinish.mb'%MBname
#             #logging.info(cmd5)
#             my_file=Path('%sFinish.mb'%MBname)
#             if my_file.is_file():
#             	pass
#             else:
#             	os.system(cmd5)
#             	logging.info(cmd5)
#             print('len equal',Already,Nocodefile)
#             time.sleep(60)
#             continue
try:
    while True:
        time.sleep(59)
        #my_observer2.stop()
        #my_observer2.join()
        break
except KeyboardInterrupt:
    # my_observer.stop()
    # my_observer.join()
    my_observer2.stop()
    my_observer2.join()
    time.sleep(10)

