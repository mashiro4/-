from tkinter import messagebox
import window
import os
import tkinter as tk


def show_popup(title, message):
    """弹窗函数"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def is_null(obj, massage):
    """判断是否为空"""
    if not obj:
        show_popup("文件转换工具&Mashiro", massage)
        return True


def file_or_directory(path):
    """判断是文件还是目录"""
    try:
        if os.path.isfile(path):
            return True
        elif os.path.isdir(path):
            return False
        elif not path:
            show_popup("文件转换工具&Mashiro", "先选择文件或目录")
    except Exception as e:
        show_popup("文件转换工具&Mashiro", "文件或目录错误")


def file_in_dic(directory_path):
    """遍历目录下的文件"""
    file_list = []
    for root, directories, files in os.walk(directory_path):
        for file in files:
            if not file.startswith("~$"):
                file_path = os.path.join(root, file)
                file_list.append(file_path)
    return file_list


def file_type(file_path):
    """判断文件类型"""
    last_name = os.path.splitext(file_path)
    if last_name[1] == ".doc":
        return "doc"
    elif last_name[1] == ".docx":
        return "docx"
    elif last_name[1] == ".pdf":
        return "pdf"
    else:
        show_popup("文件转换工具&Mashiro", "转换文件的格式需仅为‘doc’、‘docx’、‘pdf’")


if __name__ == '__main__':
    window = window.Window()


