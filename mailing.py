import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import numpy as np

def send_mail(excel_root,save_root,mail_title, mail_body, mail_address, mail_key,mail_attach):

    #excel에서 이메일과 성명을 list로 불러오기
    try:
        df = pd.read_excel(excel_root)
        email_list = np.array(df['이메일'].tolist())
        name_list=np.array(df['성명'].tolist())
    except:
        return '상태 : 메일 보내기 실패(엑셀 데이터를 불러올 수 없습니다.)'

    index=-1
    sent=0

    #email 셀에서
    for i in email_list:
        index=index+1
        # email이 있으면(nan이 아니면)
        if i!='nan':
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            try:
                s.login(mail_address, mail_key)
            except :
                return '상태 : 메일 보내기 실패(메일 서버에 로그인 할 수 없습니다.)'

            msg=MIMEBase('multipart','mixed')

            cont = MIMEText(mail_body)
            msg['Subject'] = mail_title+'_'+name_list[index]
            msg['To']=email_list[index]
            msg.attach(cont)

            pathstring=save_root+'/split/'+str(index+1)+'_'+name_list[index]+'.pdf'
            path=pathstring.encode('utf-8')
            print(path)
            part=MIMEBase("application","octet-stream")
            try:
                part.set_payload(open(path,'rb').read())
            except:
                return '상태 : 메일 보내기 실패(첨부 파일을 찾을 수 없습니다.)'
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attchment; filename='+mail_attach+".pdf")

            msg.attach(part)

            # 메일 보내기
            s.sendmail(mail_address,email_list[index], msg.as_string())
            # 세션 종료
            s.quit()

            sent=sent+1
    print("성공",sent)

    return '상태 : 메일 보내기 성공!! (' + str(sent) + '건)'
