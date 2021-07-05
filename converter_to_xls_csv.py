import string
from string import *
import os
import subprocess
import re
from win32com import client as win32
import win32com
from  docx2csv import extract

#Путь подается в ковычках  'path'
#ft_cnvrt_to_xls_csv() - главная функция
#создает файлы xls, csv  по одному на таблицу

def ft_docx_to_csv_xls_converteer(path_to_docx):
		extract(filename=path_to_docx, format="csv",singlefile=False)
		extract(filename=path_to_docx, format="xls",singlefile=False)

def ft_doc_to_docx(path):
		word = win32.genchache.EnsureDispatch('Word.Application')
		doc = word.Documents.Open(path)
		doc.Activate()
		new_file = os.path.abspath(path)
		new_file = re.sub(r'\.\w+$', 'docx', new_file)
		word.ActivateDocuments.SaveAs(new_file, FileFormat=constants.wdFormatXMLDocument)
		doc.close(False)


def ft_rtf_to_docx(path):
		word = win32com.client.Dispatch("Word.Application")
		wdFormatDocumentDefault = 16
		wdHeaderFooterPrimary = 1
		doc = word.Documents.Open(path)
		new_file = os.path.abspath(path)
		new_file = re.sub(r'\.\w+$', 'docx', new_file)
		doc.SaveAs(new_file, FileFormat=wdFormatDocumentDefault)
		doc.Close()
		word.Quit()


def ft_cnvrt_to_xls_csv(path):
		if path.find(path, ".rtf") != -1:
			ft_rtf_to_docx(path)
			path = path + 'x'
			ft_cnvrt_to_xls_csv(path)
		if path.find('.doc') != -1:
			ft_doc_to_docx(path)
			path = path + 'x'
			ft_cnvrt_to_xls_csv(path)
		if path.find('.docx') != -1:
			ft_cnvrt_to_xls_csv(path)