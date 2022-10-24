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
        if pre_col == "A":
            Class.objects.update_or_create(
                class_key="A2",
                depth=depth,
                ancestor=0,
                name=value[0],
            )
            depth += 1
            continue
        if pre_col == "M":
            break
        anc_col = chr(ord(pre_col) - 1)
        for v in value:
            if v is None:
                continue
            row = value.index(v)
            anc_name = data_json[str(anc_col)][row]
            anc_depth = depth - 1

            row = data_json[str(pre_col)].index(v)
            anc_pk = anc_col + str(row+2)
            pre_pk = pre_col + str(row+2)
            print("anc_pk ==", anc_pk)
            
            ancestor = Class.objects.get(depth=anc_depth, name=anc_name)
            print(pre_col, "==", v, "=>", pre_pk)
            Class.objects.update_or_create(
                class_key=pre_pk,
                depth=depth,
                ancestor=ancestor.class_key,
                name=v,
            )
        depth += 1