import sys
from PyQt5.QtWidgets import QWidget,QDesktopWidget,qApp,QApplication
from PyQt5.QtWidgets import QMessageBox, QStatusBar,QPushButton,QLineEdit,QGroupBox,QTextEdit,QLabel,QHBoxLayout,QVBoxLayout,QGridLayout
import main
import mailing

class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initUI(self):
        self.statusBar=QStatusBar()

        ####### 변수 세팅 ######
        self.hwp_root = ""
        self.excel_root = ""
        self.save_root = ""
        self.mail_title=""
        self.mail_body=""

        ###### 버튼 세팅 ########
        upload_form_btn = QPushButton('서식 가져오기') #버튼 텍스트, 부모 클래스(&Button1->단축키 alt+b)
        upload_excel_btn = QPushButton('엑셀 가져오기')
        save_folder_btn = QPushButton('저장 위치 선택')
        issue_cert_btn=QPushButton('수료증 발급하기')
        # download_form_btn=QPushButton('서식 다운로드')
        send_mail_btn = QPushButton('메일 발송')

        ###### 버튼 리스너 세팅 #######
        upload_form_btn.clicked.connect(self.request_hwp_root)
        upload_excel_btn.clicked.connect(self.request_excel_root)
        save_folder_btn.clicked.connect(self.request_save_root)
        issue_cert_btn.clicked.connect(self.issue_pdf)
        # download_form_btn.clicked.connect(self.download_form)
        send_mail_btn.clicked.connect(self.mail)

        ###### 텍스트박스 세팅 #######
        self.hwp_root_qle = QLineEdit()
        self.excel_root_qle = QLineEdit()
        self.save_root_qle = QLineEdit()
        self.mail_title_qle =QLineEdit()
        self.mail_body_qle = QTextEdit()
        self.mail_address_qle=QLineEdit()
        self.mail_key_qle=QLineEdit()
        self.mail_attach_qle = QLineEdit()

        self.hwp_root_qle.setDisabled(True)
        self.excel_root_qle.setDisabled(True)
        self.save_root_qle.setDisabled(True)

        ###### label 세팅 ######
        mail_address_label=QLabel("메일 주소(*) : ",self)
        mail_key_label = QLabel("메일 인증 키(*) : ", self)
        mail_title_label=QLabel("메일 제목 :",self)
        mail_body_label=QLabel("메일 내용 :",self)
        mail_attach_label = QLabel("첨부파일 명 :", self)

        ###### 그룹박스 세팅 #######
        hwp_form_gb=QGroupBox("서식(.hwp)")
        excel_gb = QGroupBox("엑셀(xls or xlsx)")
        save_gb = QGroupBox("저장위치(folder)")
        issue_cert_gb=QGroupBox("수료증 만들기")
        mail_edit_gb = QGroupBox("메일 편집기")
        mail_auth_gb = QGroupBox("메일 인증")  ##인증 관련

        ###### hbox 세팅 #######
        hwp_hbox = QHBoxLayout()
        excel_hbox = QHBoxLayout()
        save_hbox = QHBoxLayout()

        ###### vbox 세팅 #######
        self.issue_cert_vbox=QVBoxLayout()
        issuevbox=QVBoxLayout()

        ###### gridbox 세팅 #######
        mail_auth_gbox=QGridLayout()
        mail_grid_box=QGridLayout()

        hwp_hbox.addWidget(upload_form_btn)
        hwp_hbox.addWidget(self.hwp_root_qle)
        hwp_form_gb.setLayout(hwp_hbox)
        self.issue_cert_vbox.addWidget(hwp_form_gb)

        excel_hbox.addWidget(upload_excel_btn)
        excel_hbox.addWidget(self.excel_root_qle)
        # excel_hbox.addWidget(download_form_btn)
        excel_gb.setLayout(excel_hbox)
        self.issue_cert_vbox.addWidget(excel_gb)

        save_hbox.addWidget(save_folder_btn)
        save_hbox.addWidget(self.save_root_qle)
        save_gb.setLayout(save_hbox)
        self.issue_cert_vbox.addWidget(save_gb)

        self.issue_cert_vbox.addWidget(issue_cert_btn)

        issue_cert_gb.setLayout(self.issue_cert_vbox)

        mail_auth_gbox.addWidget(mail_address_label,0,0)
        mail_auth_gbox.addWidget(self.mail_address_qle,0,1)
        mail_auth_gbox.addWidget(mail_key_label,1,0)
        mail_auth_gbox.addWidget(self.mail_key_qle,1,1)

        mail_auth_gb.setLayout(mail_auth_gbox)

        mail_grid_box.addWidget(mail_title_label,0,0)
        mail_grid_box.addWidget(self.mail_title_qle,0,1)
        mail_grid_box.addWidget(mail_body_label,1,0)
        mail_grid_box.addWidget(self.mail_body_qle,1,1)
        mail_grid_box.addWidget(mail_attach_label,2,0)
        mail_grid_box.addWidget(self.mail_attach_qle,2,1)
        mail_grid_box.addWidget(send_mail_btn, 3, 1)

        mail_edit_gb.setLayout(mail_grid_box)

        issuevbox.addWidget(issue_cert_gb)
        issuevbox.addWidget(mail_auth_gb)
        issuevbox.addWidget(mail_edit_gb)
        issuevbox.addWidget(self.statusBar)

        self.setLayout(issuevbox)
        self.setWindowTitle('수료증 발급 시스템')
        self.resize(500, 300)
        self.center()
        self.show()

    def request_hwp_root(self):
        self.hwp_root=main.get_hwp_root()
        if self.hwp_root[-3:] =="hwp" or self.hwp_root[-3:]=='':
            self.hwp_root_qle.setText(self.hwp_root)
        else:
            self.warning_prompt("한글 파일(.hwp)로 작성된 서식을 사용해 주세요.")
            self.hwp_root_qle.setText('')

    def request_excel_root(self):
        self.excel_root=main.get_excel_root()
        #print(self.excel_root[-3:])
        if self.excel_root[-3:] =="xls" or self.excel_root[-3:]=='lxs' or self.excel_root[-3:]=='':
            self.excel_root_qle.setText(self.excel_root)
        else:
            self.warning_prompt("엑셀 파일(xls or xlsx)을 업로드 해 주세요.")
            self.excel_root_qle.setText('')

    def request_save_root(self):
        self.save_root=main.get_save_root()
        self.save_root_qle.setText(self.save_root)

    # 수료증 만들기(hwp->pdf->쪼개기)
    def issue_pdf(self):
        if self.hwp_root == "":
            self.warning_prompt("서식이 없습니다.")
            return
        elif self.excel_root == "":
            self.warning_prompt("엑셀 파일이 없습니다.")
            return
        elif self.save_root == "":
            self.warning_prompt("저장위치를 선택해 주세요.")
            return

        self.statusBar.showMessage('상태 : 수료증 제작 중...(마우스 클릭 금지!)')
        qApp.processEvents()
        state = main.issue_pdf(self.excel_root, self.hwp_root, self.save_root)
        if state[:2] == "에러":
            self.warning_prompt(state)
        else:
            success_state='상태 : 수료증 저장 완료(저장위치 :' + state + ')'
            self.save_root_qle.setText(state)
            self.statusBar.showMessage(success_state)

    # 메일 보내기
    def mail(self):
        print(self.excel_root)
        print(self.save_root)
        print( self.mail_title_qle.text())
        print(self.mail_body_qle.toPlainText())
        print(self.mail_address_qle.text())
        print(self.mail_key_qle.text())
        print(self.mail_attach_qle.text())
        if self.excel_root=="":
            self.warning_prompt("엑셀 값이 없습니다.")
        elif self.save_root=="":
            self.warning_prompt("저장 위치 값이 없습니다.")
        else:
            self.statusBar.showMessage('상태 : 메일 보내는 중...')
            qApp.processEvents()
            if self.mail_attach_qle.text()=="":
                self.mail_attach_qle.setText("attach")
            try:
                self.save_root=self.save_root_qle.text()
                success=mailing.send_mail(self.excel_root,self.save_root, self.mail_title_qle.text(),self.mail_body_qle.toPlainText(),self.mail_address_qle.text(),self.mail_key_qle.text(),self.mail_attach_qle.text())
                self.statusBar.showMessage(str(success))
            except:
                self.statusBar.showMessage(str(success))

    # def download_form(self):
    #     try:
    #         root=main.get_excel_form()
    #         self.statusBar.showMessage('상태 : form.xls 다운로드 완료(저장위치 :'+root+')')
    #     except:
    #         self.statusBar.showMessage('상태 : 다운로드 실패')

    def warning_prompt(self,text):
        buttonReply = QMessageBox.warning(
            self, 'warning', text,
            QMessageBox.Yes
        )
        self.statusBar.showMessage(text)

if __name__ == '__main__':
   global hwp_root, excel_root, save_root

   app = QApplication(sys.argv)
   ex = MyApp()
   sys.exit(app.exec_())