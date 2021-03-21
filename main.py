import shutil  # 파일복사용 모듈
import win32com.client as win32  # 한/글 열기 위한 모듈
import pandas as pd
from datetime import datetime
import os
import win32gui
import pdf_cut
from tkinter import Tk
import tkinter.filedialog
import pyautogui

global progress

## 수료증 한글파일로 만들기
def make_cert_hwp(excel_root,hwp_root,save_folder):
    try:
        excel = pd.read_excel(excel_root)  # 엑셀로 데이터프레임 생성
    except:
        print("error(not exists excel file)")
        return "에러 : 엑셀 파일이 존재하지 않습니다."

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 열기
        hwp.RegisterModule('FilePathCheckDLL', 'AutomationModule')  # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)
    except:
        print("error(gen_py error)")
        hwp.Quit()
        return "에러 : gen_py 에러입니다. 관리자에게 문의하세요."
    try:
        shutil.copyfile(hwp_root,  # 원본은 그대로 두고,
                        save_folder+"/certificate.hwp")  # 복사한 파일을 수정하려고 함.
    except:
        print("error(not exists hwp file)")
        hwp.Quit()
        return "에러 : 한글 파일이 존재하지 않습니다."

    try:
        hwp.Open( save_folder+"/certificate.hwp")  # 수정할 한/글 파일 열기 r"C:\python_project\reply_certi\resource\issue\certificate.hwp"
    except:
        print("error(can't open file)")
        hwp.Quit()
        return "에러 : 한글 파일을 열 수 없습니다."
    try:
        field_list = [i for i in hwp.GetFieldList().split("\x02")]  # 한/글 안의 누름틀 목록 불러오기
    except:
        hwp.Quit()
        return "에러 : 한글 파일의 필드를 불러올 수 없습니다."

    hwp.MovePos(0) # 문서 처음으로 이동
    hwp.Run('SelectAll')  # Ctrl-A (전체선택)
    hwp.Run('Copy')  # Ctrl-C (복사)
    hwp.MovePos(3)  # 문서 끝으로 이동

    try:
        for i in range(len(excel) - 1):  # 엑셀파일 행갯수-1 만큼 한/글 페이지를 복사(기존에 한쪽이 있으니까)
            hwp.Run('Paste')  # Ctrl-V (붙여넣기)
            hwp.MovePos(3)  # 문서 끝으로 이동
            pyautogui.keyDown('ctrl')
            pyautogui.press('enter')
            pyautogui.keyUp('ctrl')

        print(f'{len(excel)}페이지 복사를 완료하였습니다.')
    except:
        hwp.Quit()
        return "에러 : 한글 페이지 복사에 실패하였습니다."

    try:
        for page in range(len(excel)):  # 한/글 모든 페이지를 전부 순회하면서,
            for field in field_list:  # 모든 누름틀에 각각,
                hwp.MoveToField(f'{field}{{{{{page}}}}}')  # 커서를 해당 누름틀로 이동(작성과정을 지켜보기 위함. 없어도 무관)
                hwp.PutFieldText(f'{field}{{{{{page}}}}}',  # f"{{{{{page}}}}}"는 "{{1}}"로 입력된다. {를 출력하려면 {{를 입력.
                                 excel[field].iloc[page])  # hwp.PutFieldText("index{{1}}") 식으로 실행될 것.
    except:
        return "에러 : 한글 또는 엑셀 파일이 잘못되었습니다."

    hwp.Save()  # 한/글 파일저장
    hwp.Quit()  # 한/글 종료.

    del hwp
    return 1

## 수료증 pdf 파일로 만들기
def make_cert_pdf(save_folder):
    try:
        os.chdir(save_folder)
        hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
        hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기

        win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 숨겨줘. 0은 숨기기, 5는 보이기, 3은 풀스크린 등
        hwp.RegisterModule('FilePathCheckDLL', 'AutomationModule')  # 보안모듈 적용

        hwp.Open(save_folder+"\certificate.hwp")  # 한/글로 열어서
        hwp.HAction.GetDefault('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # PDF로 저장할 건데, 설정값은 아래와 같이.
        hwp.HParameterSet.HFileOpenSave.filename =save_folder+"\certificate.pdf"   # 확장자는 .pdf로,
        hwp.HParameterSet.HFileOpenSave.Format = 'PDF'  # 포맷은 PDF로,
        hwp.HAction.Execute('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # 위 설정값으로 실행해줘.

        win32gui.ShowWindow(hwnd, 5)  # 다시 숨겼던 한/글 창을 보여주고,
        hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
        hwp.Quit()  # 한/글 종료
        del hwp
    except:
        hwp.Quit()
        return "에러 : PDF로 변환에 실패하였습니다."
    return 1

#엑셀 파일 경로 묻기
def get_excel_root():
    root = Tk().withdraw()
    excel_root = tkinter.filedialog.askopenfilename(initialdir="/", title="엑셀 파일 업로드", filetypes={("all files", "*.*")})
    print(excel_root)
    return excel_root

# 서식(hwp) 파일 경로 묻기
def get_hwp_root():
    root = Tk().withdraw()
    hwp_root = tkinter.filedialog.askopenfilename(initialdir="/", title="서식 업로드(hwp)",
                                                  filetypes={("hwp files", "*.hwp")})
    print(hwp_root)
    return hwp_root

# 저장 폴더 경로 묻기
def get_save_root():
    root = Tk().withdraw()
    save_folder = tkinter.filedialog.askdirectory(title="저장 폴더 선택");
    print(save_folder)
    return save_folder

def get_excel_form():
    download_folder=get_save_root()
    shutil.copyfile("./resource/form.xls",  # 원본은 그대로 두고,
                    download_folder + "/form.xls")  # 복사한 파일을 수정하려고 함.
    return download_folder

def issue_pdf(excel_root,hwp_root,save_root):
    try:
        dirname=str(datetime.now().year)+"-"+str(datetime.now().month)+"-"+str(datetime.now().day)+" "+str(datetime.now().hour)+"h"+str(datetime.now().minute)+"m"+str(datetime.now().second)+"s"
        save_root = save_root +"/"+ str(dirname)
        if not os.path.exists(save_root):
            os.makedirs(save_root)
    except OSError:
        print('Error: Creating directory. ' +save_root)

    msg=make_cert_hwp(excel_root,hwp_root,save_root)

    if msg==1:
        msg=make_cert_pdf(save_root)
        if msg==1:
            msg = pdf_cut.setUp(save_root, excel_root)
            if msg==1:
                msg=pdf_cut.split()
            else:
                return msg
        else:
            return msg
    else:
        return msg

    return msg
    #return '상태 : 수료증 저장 완료(저장위치 :' + msg + ')'

# if __name__ == "__main__":
#     global excel_root, hwp_root, save_root
#     excel_root=get_excel_root()
#     hwp_root=get_hwp_root()
#     save_root=get_save_root()
#     issue_pdf(excel_root,hwp_root,save_root)

