#import pyi_splash
#pyi_splash.close()
#위 코드는 pyinstaller로 build시 주석해제
#pyinstaller xlwings_unmerge.py --splash "./splash.png" -F -w

import xlwings as xw
from tkinter import filedialog
from tkinter import messagebox
from tqdm.tk import tqdm_tk
import os

def loadfile():
    """파일을 불러오는 함수

    Returns:
        파일경로 및 파일명
    """
    files = filedialog.askopenfilenames(initialdir="/", title="파일을 선택 해 주세요", filetypes=(("*.xlsx", "*xlsx"), ("*.xls", "*xls"), ("*.xlsm", "*xlsm")))
    if not files:
        return ''
    else:
        return files[0]


print('파일을 선택창을 띄웁니다.')
loadedfile = loadfile()
if not loadedfile:
    messagebox.showwarning("경고", "파일을 추가하고 다시 실행하세요")    #파일 선택 안했을 때 메세지 출력

else:
    print("작업파일 : " + loadedfile)
    App = xw.App(visible=False)
    wb = xw.Book(loadedfile)
    wb.activate()
    for sheet in tqdm_tk(wb.sheets, desc="전체 진행률"):
        rngAll = sheet.used_range
        for rngC in tqdm_tk(rngAll, desc=sheet.name + " 작업중"):
            if rngC.merge_cells:
                val = rngC.value
                workrange = rngC.merge_area
                rngC.unmerge()
                workrange.value = val
    wb.save(os.path.join(os.path.dirname(loadedfile), 'Unmerged.xlsx'))
    messagebox.showinfo("완료", "원본폴더에 Unmerged.xlsx로 저장했습니다.")
    wb.close()
    App.kill()
