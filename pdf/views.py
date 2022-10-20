import os
import win32com.client as win32
from upload.models import Data
import json
import re

# Create your views here.

def get_excel_cell(col, row):
    data = Data.objects.get(pk=4)
    data_json = json.loads(data.data)
    # print(col+str(row))
    cell = data_json.get(col)[row-1]
    print(cell)
    return cell

def create_report():

    # mapping_string = ["U3", "U4", "AA17"]

    # for ms in mapping_string:
    #     print(ms)
    #     col_arr = re.findall('[a-zA-Z]', ms)
    #     row_arr = re.findall('\d', ms)

    #     col = ' '.join(c for c in col_arr)
    #     row = ' '.join(r for r in row_arr)

    #     print(col, row)

    # return

    # hwp 파일 열기
    file_root = os.path.abspath(os.path.join(
        os.path.dirname(__file__),
        "file"
    ))

    file_path = file_root + "/report_template.hwp"

    hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(file_path,"HWP","forceopen:true")

    # 모두 찾아 바꾸기 세부 설정
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    option=hwp.HParameterSet.HFindReplace
    option.FindString = "U3"
    # option.ReplaceString = "YEPPI YEPPI"
    option.ReplaceString = get_excel_cell("U", 3)
    option.IgnoreMessage = 1

    # 모두 찾아 바꾸기 실행
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    # 다른 이름으로 저장
    new_file_path = file_root + "/new_report.hwp"
    hwp.SaveAs(new_file_path)

    # 닫기
    hwp.Quit()

    print("replace success")