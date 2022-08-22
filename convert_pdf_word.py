#autor: Genário Carneiro
#creditos: Genário Carneiro
#data: agosto 2022

#importando as bibliotecas
import win32com.client
import os

word=win32com.client.Dispatch('word.Application')
word.visible=1

doc_pdf='Portugues.pdf'
input_file=os.path.abspath(doc_pdf)

wb=word.Documents.Open(input_file)
output_file=os.path.abspath(doc_pdf[0:-4] + 'docx'.format())
wb.SaveAs2(output_file, FileFormat=16)
print('Documento convertido')
wb.Close()

word.Quit()