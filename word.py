from docx import Document
from spire.doc import *
from spire.doc.common import *


def docx_to_pdf(input_file_path, output_dic_path, word_object, word='word', pdf='pdf'):
    """docx转为pdf"""
    try:
        output_file_path = output_dic_path + '.' + pdf
        if not os.path.exists(os.path.split(output_file_path)[0]):
            os.makedirs(os.path.split(output_file_path)[0])
        doc = word_object.Documents.Open(input_file_path)
        doc.SaveAs(output_file_path, FileFormat=17)  # 17 表示pdf文件格式
        doc.Close()
        return True
    except Exception as e:
        return False


def doc_to_docx(input_file_path, output_dic_path, word_object, doc='doc', docx='docx'):
    """doc转docx"""
    try:
        output_file_path = output_dic_path + '.' + docx
        if not os.path.exists(os.path.split(output_file_path)[0]):
            os.makedirs(os.path.split(output_file_path)[0])
        doc = word_object.Documents.Open(input_file_path)
        doc.SaveAs(output_file_path, FileFormat=16)  # 16 表示docx文件格式
        doc.Close()
        return True
    except Exception as e:
        return False


def word_add_watermark(input_file_path, output_dic_path, watermark_text, format, size, docx='docx'):
    """word加水印"""
    try:
        document = Document()
        document.LoadFromFile(input_file_path)
        txt_watermark = TextWatermark()
        # 设置文本水印的格式
        txt_watermark.Text = watermark_text
        txt_watermark.FontSize = int(size)
        txt_watermark.Color = Color.get_Gray()
        if format == "水平":
            txt_watermark.Layout = WatermarkLayout.Horizontal
        elif format == "斜体":
            txt_watermark.Layout = WatermarkLayout.Diagonal
        txt_watermark.fontName = "Arial"
        # 将文本水印添加到文档中
        document.Watermark = txt_watermark
        # 保存结果文档
        output_file_path = output_dic_path + '.' + docx
        if not os.path.exists(os.path.split(output_file_path)[0]):
            os.makedirs(os.path.split(output_file_path)[0])
        document.SaveToFile(output_file_path, FileFormat.Docx)
        document.Close()
        return True
    except Exception as e:
        return False