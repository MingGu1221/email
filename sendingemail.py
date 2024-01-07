import pandas as pd
import smtplib
from email.message import EmailMessage

# 读取文件和数据
doc_READ = pd.read_excel('C:\\Users\Ming\python\pythonporject1\email\list.xlsx')
doc_READ['gender'] = doc_READ['gender'].replace({'man': 'Mr', 'female': 'Mrs'})

def send_email(to_email, subject, body):
    sender_email = 'mwsgret@gmail.com'  # 你的Gmail地址
    app_password = 'teeg rerg etes gfds'  # 生成的16位应用专用密码

    smtp_server = 'smtp.gmail.com'
    port = 587

    msg = EmailMessage()
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = to_email

    server = smtplib.SMTP(smtp_server, port)
    server.starttls()
    server.login(sender_email, app_password)  # 使用app_password进行登录
    server.send_message(msg)
    server.quit()

# 准备邮件文本并发送
for index, row in doc_READ.iterrows():
    title = row['gender']  # 假设第一列包含称谓
    surname = row['surname']  # 假设第二列包含姓氏
    email = row['email']  # 使用实际的电子邮件列名称

    subject = f"Happy New Year {title} {surname}"
    body = f"Dear {title} {surname}:\nHappy new year. \nBest wishes,\nCris"
    # print(body)  # 打印邮件内容以供检查
    send_email(email, subject, body)
