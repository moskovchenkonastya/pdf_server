#!/usr/bin/env python
# -*- coding: utf-8 -*-
#

import string, cgi, time
from os import curdir, sep, remove
from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer
import qrcode

import urllib
from urlparse import parse_qs

from baseimageuwc import  UWCImage
from baseimagetvsz import  TVSZImage


from wand.image import Image as WandImage

# определем
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont   


from PIL import ImageFont, ImageDraw, Image as Image

# библиотека для работы с COM объектами
import win32com.client as win32


import tempfile


__CACHE__ = dict()

def need_to_convert (type_of_file):
    """Определяем требуется ли конвертация файла в формат PDF"""
    return not type_of_file == "pdf"


def convert_to_pdf(file_path, type_of_file, to_pdf_path, custom_page = None):
    """
    Функция конвертирует файл WORD, Excel в 
    PDF файл, возращает количество файлов

    """
    page_count = 0
    if type_of_file in ["doc","docx","rtf"]:
        page_count = convert_word_to_pdf (file_path, to_pdf_path, custom_page)

    elif type_of_file in ["xls","xlsx"]:
        page_count = convert_excel_to_pdf (file_path, to_pdf_path, custom_page)

    return page_count



def convert_pdf_pdf(source_pdf_file, to_pdf_path, custom_page = None):

    filedocs = PdfFileReader(file(source_pdf_file, "rb"), strict = False)
    output = PdfFileWriter()

    result = True

    page_count = filedocs.getNumPages()

    for i in range(0,page_count):

        if custom_page is None:
            page = filedocs.getPage(i)
            output.addPage(page)
        elif custom_page == i + 1:
            page = filedocs.getPage(i)
            output.addPage(page)

    res = file(to_pdf_path,"wb")
    output.write(res)

    res.close()

    return page_count

def convert_word_to_pdf(doc_path, to_pdf_path, custom_page = None):
    """
    Функция возвращает количество страниц в файле Word
    """
    
    page_count = 0
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(doc_path)
        try:
            doc.ActiveWindow.View.RevisionsView = 0                   #; //0 = wdRevisionsViewFinal
            doc.ActiveWindow.View.ShowRevisionsAndComments = False       #
            
            page_count = doc.ComputeStatistics (2) # wdStatisticPages = 2
        except Exception, e:
            pass
        if custom_page == None:
            doc.SaveAs2(to_pdf_path, 17)
        else:
#            doc.ExportAsFixedFormat(0, to_pdf_path, 0, 0, 3,custom_page,custom_page)
            doc.ExportAsFixedFormat(to_pdf_path, 17, 0, 1, 3, custom_page, custom_page)

#        doc.SaveAs2(to_pdf_path, 17)                                #; //сохраняем в pdf
        word.Application.Quit(0)

    except Exception, e:
        raise e

    return page_count

def convert_excel_to_pdf(excel_path, to_pdf_path, custom_page = None):

    # для Excel всегда будет одна страница
    page_count = 1
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        doc = excel.Workbooks.Open(excel_path)

        if custom_page == None:
            doc.ExportAsFixedFormat(0, to_pdf_path, 0)
        else:
            doc.ExportAsFixedFormat(0, to_pdf_path, 0, 0, 0, custom_page, custom_page)

        doc.Close()
        excel.Quit()

        #del excel
    except Exception, e:
        raise e

    return page_count


def convert_pdf_to_png(from_file, to_file , num_page = 1):
    ret = True

    with WandImage(filename=from_file, resolution=150) as pdf:    

        pages = len(pdf.sequence)

        num_page = max (num_page - 1, 0)
        
        if (pages <= num_page):
            ret = False
            return ret

        image = WandImage(
            width=pdf.width,
            height=pdf.height
        )

        image.composite(
            pdf.sequence[num_page],
            top=0,
            left=0
        )

        image.save(filename=to_file)

    return ret


def get_qr_code (text_data, company, reg_number, to_image_path):

    try:
        qr = qrcode.QRCode(
                    version=2,
                    error_correction=qrcode.constants.ERROR_CORRECT_L,
                    box_size=20,
                    border=1,
                )

        qr.add_data(text_data)
        qr.make(fit=True)

        image_factory = None
        if u"ОВК" in company:       
            image_factory = UWCImage
#        elif u"ТВСЗ" in company:       
#            image_factory = TVSZImage

        # im contains a PIL.Image.Image object 
        im = qr.make_image(image_factory= image_factory)


        if not reg_number == "":

            fontsize       = 15
            img_fraction_width, img_fraction_height   = 1, 0.08
            font_family     = "verdana.ttf"

            font = ImageFont.truetype(font_family, fontsize)
            while font.getsize(reg_number)[0] < img_fraction_width*im.size[0] and font.getsize(reg_number)[1] < img_fraction_height*im.size[1]:
                # iterate until the text size is just larger than the criteria
                fontsize += 5
                font = ImageFont.truetype(font_family, fontsize)            

            # optionally de-increment to be sure it is less than criteria
            fontsize -= 5

            font = ImageFont.truetype(font_family, fontsize)           

            textwidth, textheight = font.getsize(reg_number)
            basewidth, baseheight = im.size
            im.size = (basewidth, baseheight + textheight)

            newim = Image.new (im.mode, (basewidth, baseheight + textheight - 15),(255, 255, 255))
            draw = ImageDraw.Draw(newim)
            draw.text((20, baseheight - 15), reg_number, font=font)

            newim = newim.rotate(90,expand=1)            
            newim.paste (im, box =(0,0,basewidth, baseheight))


            #im = im.transform ((basewidth, baseheight + textheight), Image.EXTENT, None)

            im = newim


        im.save(to_image_path)

    except Exception, e:
        raise e

    return True


def add_qr_code (source_pdf_file, to_pdf_path, image_file):

    filedocs = PdfFileReader(file(source_pdf_file, "rb"))
    output = PdfFileWriter()

    temp_pdf_image = tempfile.TemporaryFile(mode ="w", suffix = ".pdf").name

    for i in range(0,filedocs.getNumPages()):

        page = filedocs.getPage(i)

        w,h = int(page.mediaBox[2]), int(page.mediaBox[3])

        imgDoc = canvas.Canvas(temp_pdf_image, pagesize=(w, h))

        imgDoc.drawImage(image_file, int(w) -  23*mm , int(h) - 23*mm, 20*mm, 20*mm)    
        imgDoc.save()

        overlay  = PdfFileReader(file(temp_pdf_image,"rb")).getPage(0)
        page.mergePage(overlay)

        output.addPage(page)

    res = file(to_pdf_path,"wb")
    output.write(res)

    res.close()


def get_conf(to_file, page_size, company_reg):

    c = canvas.Canvas(to_file)
    c.setPageSize (page_size)

    (w,h) = page_size
   
    pdfmetrics.registerFont(TTFont('Verdana', 'Verdana.ttf'))
    c.setFont("Verdana", 12)
    
    from reportlab.lib.colors import red

    border_size_x, border_size_y  = int(55 *mm), int(20*mm)
    padding_size_x,padding_size_y = int(55 *mm), int(20*mm)

    delta_from = int (10 * mm)

    c.setStrokeColor(red)
    c.setFillColor(red)

    c.rect (  w-border_size_x - delta_from, 0 + delta_from, border_size_x, border_size_y)

    x_vert = w - (border_size_x + delta_from )  / 2  - int(4 *mm)
    y_vert = 0 + delta_from 


    text = u"Конфиденциально"
    c.drawCentredString (x_vert, y_vert + int(13 *mm), text)

    text = u"%s" % company_reg
    c.drawCentredString (x_vert, y_vert + int(7 *mm),  u"%s" % company_reg)

    c.save()


    return True

def get_stamp(to_file, page_size, type_reg, number_reg, date_reg):

    c = canvas.Canvas(to_file)
    c.setPageSize (page_size)

    (w,h) = page_size
   
    pdfmetrics.registerFont(TTFont('Verdana', 'Verdana.ttf'))
    c.setFont("Verdana", 12)
    
    from reportlab.lib.colors import blue
    border_size_x, border_size_y  = int(55 *mm), int(20*mm)
    padding_size_x,padding_size_y = int(55 *mm), int(20*mm)

    delta_from = int (10 * mm)

    c.setStrokeColor(blue)
    c.setFillColor(blue)

    c.rect (  w-border_size_x - delta_from, 0 + delta_from, border_size_x, border_size_y)

    c.drawString (w - border_size_x - delta_from + int(4 *mm), 0 + delta_from + int(13 *mm), u"%s №%s" % (type_reg, number_reg))
    c.drawString (w - border_size_x - delta_from + int(4 *mm), 0 + delta_from + int(7 *mm),  u"%s" % date_reg)
    
    c.save()


    return True

def add_stamp (source_pdf_file, to_pdf_path, type_reg, number_reg, date_reg):

    filedocs = PdfFileReader(file(source_pdf_file, "rb"))
    output   = PdfFileWriter()

    stamp = tempfile.TemporaryFile(mode ="w", suffix = ".pdf").name
    for i in range(0,filedocs.getNumPages()):
        page = filedocs.getPage(i)

        # вставляем штамп
        if (i == 0 and get_stamp (stamp, (int(page.mediaBox[2]), int(page.mediaBox[3])), type_reg, number_reg, date_reg)) :
            overlay  = PdfFileReader(file(stamp,"rb")).getPage(0)
            page.mergePage(overlay)

        output.addPage(page)


    res = file(to_pdf_path,"wb")
    output.write(res)

    res.close()

def add_conf (source_pdf_file, to_pdf_path, company_reg):

    filedocs = PdfFileReader(file(source_pdf_file, "rb"))
    output   = PdfFileWriter()

    stamp = tempfile.TemporaryFile(mode ="w", suffix = ".pdf").name
    for i in range(0,filedocs.getNumPages()):
        page = filedocs.getPage(i)

        # вставляем штамп
        if (i == 0 and get_conf (stamp, (int(page.mediaBox[2]), int(page.mediaBox[3])), company_reg)) :
            overlay  = PdfFileReader(file(stamp,"rb")).getPage(0)
            page.mergePage(overlay)

        output.addPage(page)


    res = file(to_pdf_path,"wb")
    output.write(res)

    res.close()
    pass



class MyHandler (BaseHTTPRequestHandler):

    def do_GET(self):
        self.send_error(404, 'File Not Found: %s' % self.path)

    def save_file(self, to_file):
        
        # cчитываем размер файла
        request_sz = int(self.headers["Content-length"])
        # считываем поток
        request_str = self.rfile.read(request_sz)
        
        f = open(to_file, "wb")

        f.write(request_str)

        f.close()    

    def upload_file(self, from_file):
        self.send_response(200)
        self.send_header('Content-type',        'application/pdf')

        self.end_headers()
        f = file(from_file,"rb")
        self.wfile.write(f.read())
        f.close()    

    def upload_image(self, from_db):
        self.send_response(200)
        self.send_header('Content-type',        'image/png')
        self.send_header('X_Pages_count',        from_db["pageCount"])

        self.end_headers()

        f = file(from_db["fileName"],"rb")
        
        self.wfile.write(f.read())
        f.close()  

    def upload_qr_code(self, from_file):
        self.send_response(200)
        self.send_header('Content-type',        'image/png')
        self.end_headers()

        f = file(from_file,"rb")
        
        self.wfile.write(f.read())
        f.close()  


    def upload_error(self, from_db):
        self.send_response(404)
        self.send_header('Content-type',        'text/html')
        self.send_header('X_Pages_count',        from_db["pageCount"])
        self.end_headers()

        self.wfile.write("<html><head><title>Page not found.</title></head>")
        self.wfile.write("<body><p>Page not found.</p>")
          # If someone went to "http://something.somewhere.net/foo/bar/",
          # then self.path equals "/foo/bar/".
        self.wfile.write("</body></html>")



    def do_POST_QR(self):
        # считываем параметры из браузера
        params = parse_qs(self.path.split('?')[1])
        
        # описываем имена временных файлов
        file_type = params[u"type"][0]
        filename = tempfile.TemporaryFile(mode ="r", suffix = ("." + file_type)).name 
        pdf_result = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name 

        # сохраняем файл с сервера
        self.save_file (filename)
        # проверяем нужно ли конвертировать файл
        pdf = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name
        if (need_to_convert (file_type)):
            convert_to_pdf(filename, file_type, pdf)

        else:
            convert_pdf_pdf(filename, pdf)


        # формируем временный файл c QR кодом
        qr_code_image = tempfile.TemporaryFile(mode ="r", suffix = ".png").name

        
        text_qr         = urllib.unquote(params[u"text"][0]).decode('utf8')         if u"text" in params else "" 
        company_qr      = urllib.unquote(params[u"company"][0]).decode('utf8')      if u"company" in params else "" 
        reg_number      = urllib.unquote(params[u"reg_number"][0]).decode('utf8')   if u"reg_number" in params else ""
        
        # формирование qr - кода
        if get_qr_code (text_qr, company_qr, reg_number, qr_code_image):
            add_qr_code (pdf, pdf_result, qr_code_image)

        self.upload_file (pdf_result)

        pass


    def do_GET_QR(self):
        params = parse_qs(self.path.split('?')[1])

        file_type = params[u"type"][0]
        filename = tempfile.TemporaryFile(mode ="r", suffix = ("." + file_type)).name 

        # формируем временный файл c QR кодом
        qr_code_image = tempfile.TemporaryFile(mode ="r", suffix = ".png").name


        # формируем временный файл c QR кодом
        text_qr         = urllib.unquote(params[u"text"][0]).decode('utf8')         if u"text"       in params else "" 
        company_qr      = urllib.unquote(params[u"company"][0]).decode('utf8')      if u"company"    in params else "" 

        reg_number      = urllib.unquote(params[u"reg_number"][0]).decode('utf8')   if u"reg_number" in params else ""

        if get_qr_code (text_qr, company_qr, reg_number, qr_code_image):
            self.upload_qr_code(qr_code_image)

        #   self.upload_file (pdf_result)

        pass
    
    # вставка штампа на первую страницу
    def do_POST_STAMP(self):
        
        params = parse_qs(self.path.split('?')[1])
      

        file_type = params[u"type"][0]

        filename = tempfile.TemporaryFile(mode ="r", suffix = ("." + file_type)).name 
        pdf_result = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name 

        # сохраняем файл с сервера
        self.save_file (filename)

        pdf = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name
        if (need_to_convert (file_type)):
            convert_to_pdf(filename, file_type, pdf)

        else:
            convert_pdf_pdf(filename, pdf)

        # формируем временный файл cо штампом

        document    = urllib.unquote(params[u"document"][0]).decode('utf8') 
        reg_number  = urllib.unquote(params[u"reg_number"][0]).decode('utf8') 
        date_number = urllib.unquote(params[u"date_number"][0]).decode('utf8') 

        add_stamp (pdf, pdf_result, document,reg_number, date_number)


        self.upload_file (pdf_result) 
        pass

    
    def do_POST_CONF(self):
        params = parse_qs(self.path.split('?')[1])
      

        file_type = params[u"type"][0]

        filename = tempfile.TemporaryFile(mode ="r", suffix = ("." + file_type)).name 
        pdf_result = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name 

        # сохраняем файл с сервера
        self.save_file (filename)

        pdf = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name
        if (need_to_convert (file_type)):
            convert_to_pdf(filename, file_type, pdf)

        else:
            convert_pdf_pdf(filename, pdf)

        # формируем временный файл c QR кодом

        #document    = urllib.unquote(params[u"document"][0]).decode('utf8') 
        company     = urllib.unquote(params[u"company"][0]).decode('utf8') 

        add_conf (pdf, pdf_result, company)


        self.upload_file (pdf_result) 
        pass

    
    def do_POST_VIEWER(self):

        params = parse_qs(self.path.split('?')[1])

        file_type   = params[u"type"][0]
        num_page    = int(params[u"numpage"][0])
        filename    = tempfile.TemporaryFile(mode ="r", suffix = ("." + file_type)).name 

        key_cache   = params[u"version"][0] + "_" + str(num_page)


        # сохраняем файл с сервера
        self.save_file (filename)

        if key_cache in __CACHE__:
            self.upload_image(__CACHE__[key_cache])
        else:
            pdf = tempfile.TemporaryFile(mode ="r", suffix = ".pdf").name

            page_count = 0
            if need_to_convert(file_type):
                page_count = convert_to_pdf(filename, file_type, pdf, num_page)
            else:   
                page_count = convert_pdf_pdf(filename, pdf, num_page)

            # конвертируем файл
            png_result = tempfile.TemporaryFile(mode ="r", suffix = ".png").name

            cache_value = dict()
            cache_value["pageCount"]    = page_count
            cache_value["fileName"]     = png_result
            __CACHE__[key_cache]        = cache_value 

            # если конвертация прошла успешно
            if convert_pdf_to_png (pdf, png_result, 1):

                self.upload_image(cache_value)
            else:
                self.upload_error(cache_value)

        pass

    
    def do_POST(self):

        if ("/qr" in self.path):
            self.do_POST_QR()
        elif ("/stamp" in self.path):
            self.do_POST_STAMP()
        elif ("/conf" in self.path):
            self.do_POST_CONF()
        elif ("/viewer" in self.path):
            self.do_POST_VIEWER()


    def do_GET(self):

        if ("/qr" in self.path):
            self.do_GET_QR()
        elif ("/viewer" in self.path):
            self.do_POST_VIEWER()


def main():
    try:
        server = HTTPServer(('10.77.4.116', 81), MyHandler)
        print 'PDF: Started'
        server.serve_forever()
    except KeyboardInterrupt:
        print "PDF: Exit"
        server.socket.close()
    
if __name__ == '__main__':
    main()
