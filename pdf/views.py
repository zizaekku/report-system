import os
import win32com.client as win32

# Create your views here.
def create_report():

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
    option.ReplaceString = "YEPPI YEPPI"
    option.IgnoreMessage = 1

    # 모두 찾아 바꾸기 실행
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

    # 다른 이름으로 저장
    new_file_path = file_root + "/new_report.hwp"
    hwp.SaveAs(new_file_path)

    # 닫기
    hwp.Quit()

    print("replace success")
