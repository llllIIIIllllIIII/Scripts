import PyPDF2 
import os



fileNum = 0

filelist = []

def cls():
    os.system("cls" if os.name == "nt" else "clear")

while True:
    mergeFile = PyPDF2.PdfFileMerger()
    filelist.clear()

    fileNum = int(input("輸入要合併之檔案數: "))

    for i in range(0,fileNum):
        filelist.append(str(input("請輸入第"+str(i+1)+"個PDF檔名(ex:abc.pdf): ")))

    for PDFName in filelist:
        mergeFile.append(PyPDF2.PdfFileReader(PDFName, 'rb'))

    output = str(input("輸入合併後輸出之檔名(ex:Merge.pdf): "))

    mergeFile.write(output)
    os.system("pause")
    mergeFile.close()
    cls()
