print('로딩중..........')
import xlwings as xw
from tkinter import filedialog
from tkinter import messagebox
from tqdm import tqdm 
import os

list_file = []                                          #파일 목록 담을 리스트 생성
print('파일을 선택창을 띄웁니다.')
def loadingfile():
    files = filedialog.askopenfilenames(initialdir="/",title = "파일을 선택 해 주세요",filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
    if files =='':
        return ''
    else:
        return files[0]
        
loadedfile = loadingfile()
if loadedfile == '':
    messagebox.showwarning("경고", "파일을 추가하고 다시 실행하세요")    #파일 선택 안했을 때 메세지 출력

else:
    
    print("작업파일 : "+loadedfile)
    file_path = loadedfile
    wb=xw.Book(file_path)
    wb.activate
    for sheet in tqdm(wb.sheets,desc = "전체 진행률"):
        rngAll = sheet.used_range
        for rngC in tqdm(rngAll,desc = sheet.name + " 작업중"):
            if rngC.merge_cells:
                workrange = rngC.merge_area
                rngC.unmerge()
                workrange.value=rngC.value
    dirpath=os.path.dirname(loadedfile)
    wb.save(dirpath+'/Unmerged.xlsx')
    messagebox.showinfo("완료", "원본폴더에 Unmerged.xlsx로 저장했습니다.")
    wb.close()
quit()


