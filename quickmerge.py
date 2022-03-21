import PyPDF2 
import sys
import os



fileNum = 0

filelist = []

def cls():
    os.system("cls" if os.name == "nt" else "clear")

file_paths = sys.argv[1:]  # the first argument is the script itself



while True:
    mergeFile = PyPDF2.PdfFileMerger()
    filelist.clear()


    for p in file_paths:
        mergeFile.append(PyPDF2.PdfFileReader(p, 'rb'))

        

    output = str(input("輸入合併後輸出之檔名(ex:Merge.pdf): "))

    mergeFile.write(output)
    os.system("pause")
    mergeFile.close()
    cls()
