import time
import os, os.path
import win32com.client
from win32com.client import Dispatch
from PyPDF2 import PdfFileReader, PdfFileWriter

def find_file(path, search, ext):
    files = os.listdir(path)
    docm = filter(lambda x: x.endswith(ext), files)
    docm = list (docm)
    for name in docm :
        if ( name.find(search) > -1 ):
            return name

def compil_pdf_to_print ():
    pdf_writer = PdfFileWriter()
    paths = []
    doc_path = os.getcwd() + "\\Docs\\PDF\\"
    draw_path = os.getcwd() + "\\Drawing\\"
    
    file_name = find_file(doc_path, "BR", ".pdf")
    paths.append( doc_path + file_name)

    file_name = find_file(draw_path, "prod", ".pdf")
    paths.append( draw_path + file_name)

    file_name = find_file(doc_path, "REC-06C", ".pdf")
    paths.append( doc_path + file_name)

    file_name = find_file(doc_path, "REC-06D", ".pdf")
    paths.append( doc_path + file_name)

    file_name = find_file(draw_path, "OTK", ".pdf")
    paths.append( draw_path + file_name)

    
    for path in paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            # Добавить каждую страницу в объект писателя
            pdf_writer.addPage(pdf_reader.getPage(page))
    
    output = os.getcwd() + "\\test.pdf"
    with open(output, 'wb') as out:
        pdf_writer.write(out)
    return
    

def run_word_macro():
    try:
        wdFormatPDF = 17
        word = Dispatch("Word.Application")
        word.Visible = 1
        doc_path = os.getcwd() + "\\Docs"
        file_name = find_file(doc_path, "BR",'.docm')
        workbook = word.Documents.Open( doc_path + "\\" + file_name)

        if not os.path.isdir(doc_path + "\\PDF"):
            os.mkdir (doc_path + "\\PDF")
        
        """BR"""
        workbook.Application.Run("Чтение_из_Parameters.read_from_parameters_file2")
        workbook.Application.Run("Склейка_файлов.Glue_Files_P")
        workbook.Application.Run("Склейка_файлов.add_files")
        file_name= file_name.replace(".docm","")
        workbook.SaveAs(doc_path + "\\PDF\\" + file_name, FileFormat = wdFormatPDF)
        workbook.Close(0)

        """REC-06A"""
        #file_name = find_file("REC-06A")
        #workbook = word.Documents.Open(os.getcwd() + "\\" + file_name)
        #
        #workbook.Application.Run("OTK_File.read_from_parameters_file2")
        #workbook.Application.Run("OTK_File.start_tables")
        #workbook.Application.Run("OTK_File.read_from_parameters_file2")
        #
        #workbook.SaveAs(os.getcwd() + "\\RES\\" + file_name)
        #workbook.Close()
        
        """REC-06C"""
        file_name = find_file(doc_path, "REC-06C",'.docm')
        workbook = word.Documents.Open( doc_path + "\\" + file_name)
        
        workbook.Application.Run("OTK_File.read_from_parameters_file2")
        workbook.Application.Run("OTK_File.start_tables")
        workbook.Application.Run("OTK_File.read_from_parameters_file2")
        
        file_name= file_name.replace(".docm","")
        workbook.SaveAs(doc_path + "\\PDF\\" + file_name, FileFormat = wdFormatPDF)
        workbook.Close(0)
        
        """REC-06D"""
        #    read_from_parameters_file2
        file_name = find_file(doc_path, "REC-06D",'.docm')
        workbook = word.Documents.Open( doc_path + "\\" + file_name)
        
        workbook.Application.Run("Чтение_из_Parameters.read_from_parameters_file2")
        
        file_name= file_name.replace(".docm","")
        workbook.SaveAs(doc_path + "\\PDF\\" + file_name, FileFormat = wdFormatPDF)
        workbook.Close(0)

        word.Quit()
    except IOError:
        print("Error")


if __name__ == "__main__":
    
    #print (find_file("REC-06A"))
    run_word_macro()
    compil_pdf_to_print()
    
    time.sleep(2)