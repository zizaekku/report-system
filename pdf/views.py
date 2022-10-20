from http.client import ImproperConnectionState
import os
import win32com.client as win32
from upload.models import Data
import json
import re
from .models import *

# Create your views here.

def get_excel_cell(col, row):
    data = Data.objects.get(pk=4)
    data_json = json.loads(data.data)
    cell = data_json.get(col)[row-2]
    print(cell)
    return cell

def create_report():

    mapping_string = Mapping.objects.all()

    # hwp 파일 열기
    file_root = os.path.abspath(os.path.join(
        os.path.dirname(__file__),
        "file"
    ))

    file_path = file_root + "/report_template.hwp"

    hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(file_path,"HWP","forceopen:true")

    for ms in mapping_string:
        findstring = "[ " + str(ms) + " ]"
        print(findstring)

        # 모두 찾아 바꾸기 세부 설정
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
        option=hwp.HParameterSet.HFindReplace
        option.FindString = findstring
        option.ReplaceString = get_excel_cell(ms.col, ms.row)
        option.IgnoreMessage = 1

        # 모두 찾아 바꾸기 실행
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    # 다른 이름으로 저장
    u3 = Mapping.objects.get(col="U", row=3)
    u3_cell = get_excel_cell(u3.col, u3.row)
    u4 = Mapping.objects.get(col="U", row=4)
    u4_cell = get_excel_cell(u4.col, u4.row)

    new_filename = u3_cell + u4_cell + " 용역 최종보고서.hwp"
    
    new_file_path = file_root + "/" + new_filename
    hwp.SaveAs(new_file_path)

    # 닫기
    hwp.Quit()

    print("create report success")