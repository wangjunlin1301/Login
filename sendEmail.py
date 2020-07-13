import smtplib
from email.mime.text import MIMEText
from email.header import Header
from dateV import USDate
from jira import blogName
from emailContent import getEmailContent


class SendEmail():
    def __init__(self):
        #邮件服务
        self.mail = "smtpout.secureserver.net"
        self.userName = "wang.junlin@xbosoft.com"
        self.userPassd = "khalil1"

        #收件人, will modify after tested.
        self.sender = 'wang.junlin@xbosoft.com'
        self.receivers = ['wang.junlin@xbosoft.com']

        #正文
        self.message = MIMEText(getEmailContent(), 'plain', 'utf-8')
        self.message['From'] = Header(self.sender, 'utf-8')
        self.message['To'] = Header(self.receivers[0], 'utf-8')

        #邮件title,
        subject = '[TEST UPDATES, %s] %s' % (USDate, blogName)
        self.message['Subject'] = Header(subject, 'utf-8')

    def sendEmail(self):
        smtpObj = smtplib.SMTP()
        smtpObj.connect(self.mail, 25)  # 80/465/3535/25
        smtpObj.login(self.userName, self.userPassd)
        smtpObj.sendmail(self.sender, self.receivers, self.message.as_string())
        smtpObj.sendmail
        print("success!")


if __name__ == "__main__":
    mainexc = SendEmail()
    mainexc.sendEmail()