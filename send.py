import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage
import pandas as pd

def send_mail(fromUser, pwd, toUser, message, title):
    gmailUser = fromUser
    gmailPassword = pwd

    # 產生郵件資訊
    msg = MIMEText(message, "html", "utf-8")
    msg['From'] = fromUser
    msg['To'] = toUser
    msg['Subject'] = title

    # 登入Gmail郵件伺服器
    mailServer = smtplib.SMTP('smtp.gmail.com', 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(gmailUser, gmailPassword)

    # 送信
    status = mailServer.sendmail(fromUser, toUser, msg.as_string())
    if status=={}:
        print("郵件傳送成功!")
    else:
        print("郵件傳送失敗!")
    mailServer.close()

if __name__ == '__main__':
    gmailUser = ''
    gmailPassword = ''
    # 讀取 Excel 表
    data = pd.read_excel (r'list.xlsx')
    df = pd.DataFrame(data)

    # 處理每一筆資料
    for index, row in df.head().iterrows():
        _id = str(row['序號']).zfill(3);
        name = row['姓名'].strip()
        toMail = row['email'].strip()

        # 準備郵件內容
        msg = '''
        <!doctype html>
        <html>
        <head>
            <meta charset='utf-8'>
            <title>HTML mail</title>
        </head>
        <body>
          <h4>{name} 先進您好，</4>
          <p>
          感謝您參與報名「台灣民眾黨地方政治研習班台中場」活動，收到此信件代表您已報名成功。 
          </p>
          <p>
          當日請憑【報名序號{_id}】至簽到處簽名報到後入場，研習活動於B1國際演講廳，廳內禁止飲食、大聲喧嘩、拍照攝影等行為。因疫情防疫措施，敬請配戴口罩並保持安全社交距離以維護您我的健康。 
          </p>
          <p>
          再次感謝您的參與，我們6月6日不見不散。若有任何問題與建議，也歡迎隨時與我們聯繫。
          </p>
        </body>
        </html>
        '''.format(name=name, _id=_id)

        # 準備郵件標題
        title = '【報名序號{_id}】台灣民眾黨地方政治研習班台中場'.format(_id=_id)

        # 呼叫送郵件的function
        send_mail(gmailUser, gmailPassword, toMail, msg, title)
