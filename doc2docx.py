from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants

def save_as_docx(path):

    # Открываем Word
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()

        # Меняем расширение на .docx и добавляем в путь папку
        # для складывания конвертированных файлов
        new_file_abs = str(os.path.abspath(path)).split("\\")
        new_dir_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}"
        new_file_abs = f"{new_file_abs[0]}\\{new_file_abs[1]}\\doc_convert\\{new_file_abs[2]}"
        new_file_abs = os.path.abspath(new_file_abs)
        if not os.path.isdir(f'{new_dir_abs}\\doc_convert'):
            os.mkdir(f'{new_dir_abs}\\doc_convert')
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
        #print(new_file_abs)

        # Сохраняем и закрываем
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
    except:
        return str(path).split("\\")[-1]

def path_doc(paths):
    dict_error_file = []
    for path in paths:
        err = save_as_docx(path)
        if err != None:
            dict_error_file.append(err)
    if len(dict_error_file) >= 1:
        print(f'\nНе конвертированные файлы (ошибка открытия - файл поврежден):\n{dict_error_file}')

def main():
    dirs = 'C:\CHANGES\docs'
    paths = glob(f'{dirs}\\*.doc', recursive=True)
    path_doc(paths)

if __name__ == "__main__":
    main()