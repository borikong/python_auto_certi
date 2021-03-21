import configparser
import os
import PyPDF2
import pandas as pd
import numpy as np

config = configparser.ConfigParser()
config.read(os.path.dirname(os.path.realpath(__file__)) + os.sep + 'envs' + os.sep + 'property.ini')


def setUp(pdf,excel_root):
    global SAVE_ROOT, ORIGIN_FILE, NAME_DATA, ORIGIN_ROOT
    ORIGIN_ROOT=pdf
    ORIGIN_FILE = pdf+"\certificate.pdf"
    SAVE_ROOT=pdf+"/split"
    df = pd.read_excel(excel_root)
    try:
        if not os.path.exists(SAVE_ROOT):
            os.makedirs(SAVE_ROOT)
    except OSError:
        print('Error: Creating directory. ',SAVE_ROOT)

    try:
        name_list = np.array(df['성명'].tolist())
        NAME_DATA=name_list
    except:
        print("이름이 없다")
        return "에러 : 엑셀 파일이 잘못되었습니다."

    return 1


#쪼개기
def split(): #디폴트는 영어로 추출
    s_fileNum=0

    for i in NAME_DATA:
        print(">>>>>>Start Page Num : ", s_fileNum, ">>>>>>Last Page Num : ", s_fileNum + 1)
        try:
            name = NAME_DATA[s_fileNum]
        except:
            name = ""

        s_fileName = str(s_fileNum+1) + "_" + str(name) + ".pdf"
        save = SAVE_ROOT + "/"+s_fileName
        extract_tree(ORIGIN_FILE, save, s_fileNum, s_fileNum + 1)
        s_fileNum=s_fileNum+1

    return ORIGIN_ROOT

def extract_tree(in_file, out_file, start_num, last_num):
    with open(in_file, 'rb') as infp:
        # Read the document that contains the tree (in its first page)
        reader = PyPDF2.PdfFileReader(infp)
        writer = PyPDF2.PdfFileWriter()
        for i in range(start_num,last_num):
            writer.addPage(reader.getPage(i))
        with open(out_file, 'wb') as outfp:
            writer.write(outfp)
