import smtplib
from logger import logger
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from smtplib import SMTPException


#发送邮件部分
smtpserver = "smtp.exmail.qq.com"
username = "xxsjfxyj@jycinema.com"
password = "JYshuju666"
sender = "xxsjfxyj@jycinema.com"
receiver = "zb_wangxin@jycinema.com"
#抄送名单
receiver_cc = "chenchengying@jycinema.com,xieminchao@jycinema.com"
smtp = smtplib.SMTP_SSL(smtpserver,465)
smtp.ehlo()
smtp.login(user = username,password = password)

def send_mail(excel_file):
    subject = excel_file.replace(".xlsx","")

    #主题部分
    msg = MIMEMultipart()
    msg["From"] = formataddr(["信息数据分析研究中心",sender])
    msg["Subject"] = subject
    msg["To"] = receiver
    msg["Cc"] = receiver_cc
    receiver_total = [receiver]
    receiver_total.extend(receiver_cc.split(","))
    #文本部分
    text = MIMEText(subject)
    msg.attach(text)
    #excel附件部分
    att = MIMEApplication(open(excel_file,"rb").read())
    att.add_header("Content-Disposition","attachment",filename = ("GBK","",excel_file))
    msg.attach(att)
    #发送邮件
    try:
        smtp.sendmail(sender,receiver_total,msg.as_string())
        print("send mail success\n")
        logger.info("send mail success")
    except SMTPException as e:
        print("send mail failed")
        logger.info("send mail failed")
        
    smtp.quit()