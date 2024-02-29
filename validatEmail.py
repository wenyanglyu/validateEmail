import smtplib
import time
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(sender_email, sender_password, receiver_email, subject, message):
    # 设置发件人、收件人和邮件内容
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    # 连接到 Outlook 邮件服务器并发送邮件
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)  # Outlook SMTP 地址和端口号
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        return True
    except Exception as e:
        return False
    finally:
        server.quit()

def check_latest_email(mail):
    # 搜索最新的未读邮件
    result, data = mail.search(None, 'UNSEEN')

    # 如果没有未读邮件，则返回 None
    if not data[0]:
        return None

    latest_email_id = data[0].split()[-1]  # 获取最新一封未读邮件的ID

    return latest_email_id

def check_inbox(sender_email, sender_password, receiver_email):
    # 连接到收件箱的 IMAP 服务器
    mail = imaplib.IMAP4_SSL('outlook.office365.com')
    mail.login(sender_email, sender_password)
    mail.select('inbox')

    # 获取最新一封未读邮件的ID
    latest_email_id = check_latest_email(mail)

    # 如果没有未读邮件，则打印消息并退出检查
    if latest_email_id is None:
        print("收件箱中没有未读邮件。")
        return False

    # 获取最新一封未读邮件的内容
    result, data = mail.fetch(latest_email_id, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    # 检查是否为目标收件人和指定内容的邮件
    if receiver_email in msg['To'] and "The recipient's email address isn't listed in the domain's directory." in str(msg):
        return True
    return False

# 用户输入用户名
userName = input("请输入用户名：")
receiver_email = f"{userName}@autuni.ac.nz"

# 在这里填写你的邮箱账户信息和要发送的邮件内容
sender_email = "xxxx@outlook.com"
sender_password = "yyyyyy"
subject = "欢迎加入AUT 2024届2月开学（大群）"

# 创建 HTML 格式的邮件内容
# HTML格式的消息
message = """
您好，

欢迎加入AUT 2024届2月开学（大群），预祝您在AUT有个快乐的学习和生活。

如果您并未发起申请，可能您的学生用户名被冒用了，请回复此邮件，我会将该用户踢出群。

祝好，群主Yilong


"""

# 发送邮件
if send_email(sender_email, sender_password, receiver_email, subject, message):
    print(f"邮件已发送！\n正在检测{receiver_email}是否存在\n")

    # 倒计时 60 秒
    for i in range(30, 0, -1):
        print(f"检查对方是否收到倒计时: {i} 秒", end='\r')
        time.sleep(1)
    if check_inbox(sender_email, sender_password, receiver_email):
        print("\n\n邮件发送失败！\n该用户不存在！！！\n拒绝该用户入群\n")
    else:
        print("\n\n邮件对方已收到，\n该用户存在\n允许该用户入群\n")
else:
    print("邮件发送失败，请重试。")
