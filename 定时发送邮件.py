from win32file import CreateFile, SetFileTime, GetFileTime, CloseHandle
from win32file import GENERIC_READ, GENERIC_WRITE, OPEN_EXISTING
from pywintypes import Time
from threading import Timer
from email.header import Header
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import time, os, random, zipfile, smtplib
import datetime

def ModifyFileTime(filePath, lastTime):
    # 构造时间
    fh = CreateFile(filePath, GENERIC_READ | GENERIC_WRITE, 0, None, OPEN_EXISTING, 0, 0)
    offsetSec = random.randint(1200, 1500)
    cTime, aTime, wTime = GetFileTime(fh)
    aTime = Time(time.mktime(lastTime) - offsetSec)
    wTime = Time(time.mktime(lastTime) - offsetSec)
    # 修改时间属性
    SetFileTime(fh, cTime, aTime, wTime)
    CloseHandle(fh)


def GetFilePathList(filePath, result):
    if os.path.isdir(filePath):
        files = os.listdir(filePath)
        for file in files:
            if os.path.isdir(filePath + '/' + file):
                GetFilePathList(filePath + '/' + file, result)
            else:
                result.append(filePath + '/' + file)
    else:
        result.append(filePath)


def RunZipFile(filePath):
    # 构造压缩包路径
    zipFilePath = os.path.basename(filePath)
    if os.path.isdir(filePath):
        zipFilePath = zipFilePath + '.zip'
    else:
        zipFilePath = os.path.splitext(zipFilePath)[0] + '.zip'
    # 文件已存在则删除原有文件
    if os.path.exists(zipFilePath):
        os.remove(zipFilePath)
    # 创建压缩包
    zipFileList = []
    GetFilePathList(filePath, zipFileList)
    f = zipfile.ZipFile(zipFilePath, 'w', zipfile.ZIP_DEFLATED)
    for fileP in zipFileList:
        arcName = os.path.basename(fileP)
        if os.path.isdir(filePath):
            dirName = os.path.dirname(fileP)
            arcName = dirName.replace(os.path.dirname(dirName), '') + '/' + arcName
        f.write(fileP, arcName)
    f.close()
    return zipFilePath


def SendEmail(filePath):
    # 邮件设置
    fromAddr = '18852865556@163.com'  # 发件地址
    autho = '123456a'  # 授权码
    smtpServer = 'smtp.163.com'  # smtp服务器

    # 邮件信息
    toAddr = '1923088635@qq.com'  # 收件地址
    title = '昨日储值'  # 主题
    text = '昨日储值'  # 正文
    annexPath = filePath  # 附件路径

    # 构造邮件
    message = MIMEMultipart()
    message['From'] = fromAddr
    message['To'] = toAddr
    message['Subject'] = Header(title, 'utf-8')
    message.attach(MIMEText(text, 'plain', 'utf-8'))

    # 构造附件
    annex = MIMEApplication(open(annexPath, 'rb').read())
    annex.add_header('Content-Disposition', 'attachment', filename=os.path.basename(annexPath))
    message.attach(annex)

    # 发送邮件
    server = smtplib.SMTP_SSL(smtpServer, 465)
    server.login(fromAddr, autho)
    server.sendmail(fromAddr, [toAddr], message.as_string())
    server.quit()


def PreparingFilesAndSend(filePath, runTime):
    annexPath = filePath
    if os.path.isdir(filePath):
        # 遍历文件,修改时间属性
        modFileList = []
        GetFilePathList(filePath, modFileList)
        for fileP in modFileList:
            ModifyFileTime(fileP, runTime)
        # 打包
        annexPath = RunZipFile(filePath)
    else:
        # 修改时间属性
        ModifyFileTime(filePath, runTime)
        # 判断是否需要压缩
        if os.path.getsize(filePath) / float(1024 * 1024) > 10:
            annexPath = RunZipFile(filePath)
    # 发送邮件
    SendEmail(annexPath)


def TimedTask(filePath, runTime):
    interval = time.mktime(runTime) - time.mktime(time.localtime())
    Timer(interval, PreparingFilesAndSend, args=(filePath, runTime)).start()


fp = r"C:/Users/Administrator/Documents/a.xlsx"  #C:\Users\Administrator\Documents


today=datetime.date.today()
# print(today)
# rt = "2020-1-7 20:30:00"
# 转化为结构化时间
rt=str(today)+" 20:30:00"
timeFormat = "%Y-%m-%d %H:%M:%S"
s_time = time.strptime(rt, timeFormat)



if __name__ == "__main__":
    TimedTask(fp, s_time)