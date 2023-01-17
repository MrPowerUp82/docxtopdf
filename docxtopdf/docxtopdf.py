import os
import platform

def convert(docx_file: str, output_pdf: str)->None:
    if platform.system().lower() == 'windows':
        import win32com.client

        word = win32com.client.Dispatch("Word.application")

        try:
            wordDoc = word.Documents.Open(docx_file, False, False, False)
            wordDoc.SaveAs2(output_pdf, FileFormat = 17)
            wordDoc.Close()
        except Exception:
            print('Falha ao converter: {}'.format(output_pdf))

        word.Quit()

    elif platform.system().lower() == 'linux':
        os.system(f'soffice --headless --convert-to pdf {docx_file}')
        os.rename(docx_file.split('/')[-1].replace(docx_file.split('.')[-1], 'pdf'), output_pdf)

    else:
        print('Sistema desconhecido.')