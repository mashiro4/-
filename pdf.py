from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import os


def pdf_add_watermark(input_file_path, output_dic_path, pdf='pdf'):
    """pdf加水印"""
    try:
        # 读取原始PDF
        output_file_path = output_dic_path + '.' + pdf
        input_pdf = PdfReader(input_file_path)
        output_pdf = PdfWriter()
        watermark = PdfReader("watermark.pdf").pages[0]  # 使用 pages 属性获取页面
        # 为每一页添加水印
        for page_number, page in enumerate(input_pdf.pages):
            page.merge_page(watermark)
            output_pdf.add_page(page)
        # 保存结果
        if not os.path.exists(os.path.split(output_file_path)[0]):
            os.makedirs(os.path.split(output_file_path)[0])
        output_path = os.path.dirname(output_file_path)  # 获取输出PDF文件的路径
        output_filename = os.path.join(output_path, os.path.basename(output_file_path))  # 创建输出PDF文件的完整路径
        with open(output_filename, "wb") as output:
            output_pdf.write(output)
        return True
    except Exception as e:
        return False


def creat_watermark(watermark_text, format, size, font_path="simsun.ttc"):
    """
    创建水印模板
    """
    angle = 45
    if format == "水平":
        angle = 0
    elif format == "斜体":
        angle = 45
    c = canvas.Canvas("watermark.pdf", pagesize=letter)
    pdfmetrics.registerFont(TTFont('CustomFont', font_path))  # 注册自定义字体
    c.setFont("CustomFont", int(size))  # 设置字体为自定义字体，字号为80
    c.setFillGray(0.5, 0.5)  # 设置水印颜色
    c.rotate(int(angle))  # 将水印旋转45度
    c.drawString(270, 80, watermark_text)  # 设置水印位置
    c.save()
