#!/usr/bin/env python3
"""测试邮件发送"""
import smtplib
from email.mime.text import MIMEText
from email.header import Header

TO_EMAIL = "3796401811@qq.com"
FROM_EMAIL = "3796401811@qq.com"
FROM_PASSWORD = "tpugbxmctffucegj"
SMTP_HOST = "smtp.qq.com"
SMTP_PORT = 465

subject = "测试邮件 - Dragon小弟"
content = """这是测试邮件。

如果你收到这封邮件，说明币安价格提醒功能可以正常工作了。

- Dragon小弟 🫡
"""

msg = MIMEText(content, 'plain', 'utf-8')
msg['From'] = FROM_EMAIL
msg['To'] = TO_EMAIL
msg['Subject'] = Header(subject, 'utf-8')

try:
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(FROM_EMAIL, FROM_PASSWORD)
        server.sendmail(FROM_EMAIL, [TO_EMAIL], msg.as_string())
    print(f"邮件已发送到 {TO_EMAIL}")
except Exception as e:
    print(f"错误: {e}")
