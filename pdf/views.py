import os
import win32com.client as win32
from upload.models import Data
import json
from .models import *
from upload.models import *
from django.shortcuts import render
import pythoncom


# Create your views here.
def home(request):
    create_report()
    return render(request, 'home.html')

def get_excel_cell(col, row):
    data = Data.objects.get(pk=4)
    data_json = json.loads(data.data)
    cell = data_json.get(col)[row-2]
    print(cell)
    if cell == None:
        cell = "-"
    return cell

def create_report():

    pythoncom.CoInitialize()

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

        if "image" in str(option.ReplaceString):
            print("이미지")
            image = Image.objects.get(image=option.ReplaceString+".png")
            # print(image)
            hwp.InsertPicture(image.image.path, Embedded=True) # 이미지 삽입

            hwp.FindCtrl() # 이미지 선택
            hwp.HAction.Run("Cut") # 잘라내기 (커서에서 인접한 개체 선택)

            while True:

                # 이미지 삽입할 위치 찾기
                hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                hwp.HParameterSet.HFindReplace.FindString = findstring
                hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
                result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                print("=>", result)

                # 다 바꿨으면 종료
                if result == False:
                    break

                # 이미지 붙여넣기
                hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
                hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)

        else:
            print("텍스트")

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

    pythoncom.CoUninitialize()

    print("create report success")


def insert_image():

    file_root = os.path.abspath(os.path.join(
        os.path.dirname(__file__),
        "file"
    ))

    file_path = file_root + "/report_template.hwp"

    hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(file_path,"HWP","forceopen:true")

    image = Image.objects.get(pk=144)

    hwp.InsertPicture(image.image.path, Embedded=True) # 이미지 삽입
    hwp.FindCtrl() # 이미지 선택
    hwp.HAction.Run("Cut") # 잘라내기 (커서에서 인접한 개체 선택)

    while True:

        # 이미지 삽입할 위치 찾기
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        hwp.HParameterSet.HFindReplace.FindString = "[ U18 ]"
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        result = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
        print("=>", result)

        # 다 바꿨으면 종료
        if result == False:
            break

        # 이미지 붙여넣기
        hwp.HAction.GetDefault("Paste", hwp.HParameterSet.HSelectionOpt.HSet)
        hwp.HAction.Execute("Paste", hwp.HParameterSet.HSelectionOpt.HSet)


    hwp.SaveAs(file_root + "/insertimagetest.hwp")
    hwp.Quit()