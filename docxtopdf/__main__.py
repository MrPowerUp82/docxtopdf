import os
import argparse
import platform

my_parse = argparse.ArgumentParser(description="CLI e Lib.")

my_parse.add_argument("--output", action="store", type=str, help="Saída do pdf.")
my_parse.add_argument("--docx", action="store", type=str, help="Arquivo docx.")

args = my_parse.parse_args()

if args.output and args.docx:

    if platform.system().lower() == 'windows':
        import win32com.client

        word = win32com.client.Dispatch("Word.application")

        try:
            wordDoc = word.Documents.Open(args.docx, False, False, False)
            wordDoc.SaveAs2(args.output, FileFormat = 17)
            wordDoc.Close()
        except Exception:
            print('Falha ao converter: {}'.format(args.output))

        word.Quit()

    elif platform.system().lower() == 'linux':
        os.system(f'soffice --headless --convert-to pdf {args.docx}')
        os.rename(args.docx.split('/')[-1].replace(args.docx.split('.')[-1], 'pdf'), args.output)

    else:
        print('Sistema desconhecido.')