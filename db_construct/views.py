from openpyxl import load_workbook
import json

from .models import Class
from upload.models import Data


# Create your views here,
def upload():
    data = Data.objects.get(pk=4)
    data_json = json.loads(data.data)
    
    depth = 0
    for pre_col, value in data_json.items():
        print("pre_col ==", pre_col)
        
        # A 루트 생성
        if pre_col == "A":
            Class.objects.update_or_create(
                class_key="A2",
                depth=depth,
                ancestor=0,
                name=value[0],
            )
            depth += 1
            continue

        # M 이면 분기 끝내기
        if pre_col == "M":
            break

        # 조상 컬럼 구하기  ex) C라면 B, B라면 A
        anc_col = chr(ord(pre_col) - 1)
        for v in value:
            if v is None:
                continue

            # v 값이 시작하는 맨 앞 행 가져오기
            row = value.index(v)

            # 가져온 값에서 조상 cell 값 구하기
            anc_name = data_json[str(anc_col)][row]

            # 조상의 깊이 = 현재 깊이 - 1
            anc_depth = depth - 1

            # 현재 cell 값이 시작하는 행 구하기
            # row = data_json[str(pre_col)].index(v)

            # 엑셀이 2행부터 시작하므로 +2
            anc_pk = anc_col + str(row+2)
            pre_pk = pre_col + str(row+2)
            print("anc_pk ==", anc_pk)
            
            # 조상이 존재하는지 찾기
            ancestor = Class.objects.get(depth=anc_depth, name=anc_name)
            print(pre_col, "==", v, "=>", pre_pk)

            # 현재 노드 값 생성
            Class.objects.update_or_create(
                class_key=pre_pk,
                depth=depth,
                ancestor=ancestor.class_key,
                name=v,
            )
        depth += 1