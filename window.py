import os.path
import tkinter as tk
from tkinter import filedialog
import win32com.client
import main
import pdf
import word
import time


class Window:
    def __init__(self):
        # 创建窗口对象
        self.root = tk.Tk()
        self.root.title("文件转换工具&Mashiro")
        self.root.geometry("600x250")
        # 创建框架
        frame1 = tk.Frame(self.root)
        frame1.pack(fill="both", expand=True)
        frame2 = tk.Frame(self.root)
        frame2.pack(fill="both", expand=True)
        frame3 = tk.Frame(self.root)
        frame3.pack(fill="both", expand=True)
        frame4 = tk.Frame(self.root)
        frame4.pack(fill="both", expand=True)
        frame5 = tk.Frame(self.root)
        frame5.pack(fill="both", expand=True)
        """框架1"""
        # 功能页面
        self.button1_1 = tk.Button(frame1, text="选择文件", command=self.set_in_path_file)  # 选文件按钮
        self.button1_1.grid(row=0, column=0, sticky="w")
        self.button1_2 = tk.Button(frame1, text="选择目录", command=self.set_in_path_directory)  # 选目录按钮
        self.button1_2.grid(row=0, column=1, sticky="w")
        self.in_path = tk.StringVar()
        self.label1 = tk.Label(frame1, textvariable=self.in_path)
        self.label1.grid(row=1, column=0, sticky="w", columnspan=5)
        """框架2"""
        # 功能选择
        self.var = tk.StringVar()
        self.var.set('1')
        self.label2_1 = tk.Label(frame2, text="操作：")
        self.label2_1.grid(row=0, column=0, sticky="w")
        radiobutton2_1 = tk.Radiobutton(frame2, text='doc转docx', variable=self.var, value='1')
        radiobutton2_1.grid(row=0, column=1, sticky="")
        radiobutton2_2 = tk.Radiobutton(frame2, text='docx转为pdf', variable=self.var, value='2')
        radiobutton2_2.grid(row=0, column=2, sticky="")
        radiobutton2_3 = tk.Radiobutton(frame2, text='pdf加水印', variable=self.var, value='3')
        radiobutton2_3.grid(row=0, column=3, sticky="")
        radiobutton2_4 = tk.Radiobutton(frame2, text='docx加水印', variable=self.var, value='4')
        radiobutton2_4.grid(row=0, column=4, sticky="")
        """框架3"""
        # 水印参数
        self.label3 = tk.Label(frame3, text="水印文本：")  # 水印文本
        self.label3.grid(row=0, column=0, sticky="w")
        self.watermark = tk.StringVar()
        self.watermark.set("")
        self.entry1 = tk.Entry(frame3, textvariable=self.watermark, width=30)
        self.entry1.grid(row=0, column=1, sticky="w")
        self.label3 = tk.Label(frame3, text="字号：")  # 水印字号
        self.label3.grid(row=0, column=2, sticky="w")
        options1 = [36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144]
        self.size = tk.StringVar()
        self.size.set(options1[0])
        self.option_menu1 = tk.OptionMenu(frame3, self.size, *options1)
        self.option_menu1.grid(row=0, column=3, sticky="w")
        self.label4 = tk.Label(frame3, text="版式：")  # 水印版式
        self.label4.grid(row=0, column=4, sticky="w")
        options2 = ["水平", "斜式"]
        self.format = tk.StringVar()
        self.format.set(options2[1])
        self.option_menu2 = tk.OptionMenu(frame3, self.format, *options2)
        self.option_menu2.grid(row=0, column=5, sticky="w")
        """框架4"""
        self.button4_1 = tk.Button(frame4, text="输出路径", command=self.set_out_path_directory)  # 输出路径
        self.button4_1.grid(row=0, column=0, sticky="w")
        self.out_path = tk.StringVar()
        self.label4_1 = tk.Label(frame4, textvariable=self.out_path)
        self.label4_1.grid(row=1, column=0)
        """框架5"""
        self.button5_1 = tk.Button(frame5, text="开始转换", command=self.start)  # 输出路径
        self.button5_1.grid(row=0, column=0, sticky="ew")
        self.finish = tk.StringVar()
        self.label5_1 = tk.Label(frame5, textvariable=self.finish)
        self.label5_1.grid(row=0, column=1)
        # 运行主循环
        self.root.mainloop()

    def set_in_path_file(self):
        """选择输入文件"""
        file_path = filedialog.askopenfilename()
        if file_path:
            self.in_path.set(file_path)

    def set_in_path_directory(self):
        """选择输入目录"""
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.in_path.set(directory_path)

    def set_out_path_directory(self):
        """选择输出目录"""
        directory_path = filedialog.askdirectory()
        if directory_path:
            self.out_path.set(directory_path)

    def finsh_or_error(self, result):
        """是否完成"""
        if result:
            self.finish.set("转换完成，保存在：" + self.out_path.get())
        else:
            self.finish.set("转换失败")

    def file_path_handing(self):
        """文件路径处理"""
        one = self.out_path.get() + '/' + os.path.basename(self.in_path.get())
        two = os.path.splitext(one)
        return two[0]

    def dic_path_handing(self, input_file_path):
        """目录路径处理"""
        one = os.path.relpath(input_file_path, self.in_path.get())
        two = self.out_path.get() + '/' + one
        three = os.path.splitext(two)
        return three[0]

    def start(self):
        """开始转换之后"""
        # 判断输入参数是狗完整
        if main.is_null(self.in_path.get(), "请选择要转换的文件或目录"):
            return
        elif main.is_null(self.out_path.get(), "请选择输出路径"):
            return
        else:
            # 判断是文件还是目录   input_file_path  output_dic_path
            if main.file_or_directory(self.in_path.get()):  # 是文件
                # 判断操作
                if self.var.get() == '1':  # doc转docx
                    file_type = main.file_type(self.in_path.get())
                    if file_type != 'doc':
                        self.finish.set("文件类型错误，仅为doc文件")
                        return
                    word_object = win32com.client.Dispatch("Word.Application")
                    output_dic_path = Window.file_path_handing(self)
                    result = word.docx_to_pdf(self.in_path.get(), output_dic_path, word_object)
                    word_object.Quit()
                    Window.finsh_or_error(self, result)
                elif self.var.get() == '2':  # docx转pdf
                    file_type = main.file_type(self.in_path.get())
                    if file_type != 'docx':
                        self.finish.set("文件类型错误，仅为docx文件")
                        return
                    word_object = win32com.client.Dispatch("Word.Application")
                    output_dic_path = Window.file_path_handing(self)
                    result = word.doc_to_docx(self.in_path.get(), output_dic_path, word_object)
                    word_object.Quit()
                    Window.finsh_or_error(self, result)
                elif self.var.get() == '3':  # pdf加水印
                    file_type = main.file_type(self.in_path.get())
                    if file_type != 'pdf':
                        self.finish.set("文件类型错误，仅为pdf文件")
                        return
                    pdf.creat_watermark(self.watermark.get(), self.format.get(), self.size.get())
                    output_dic_path = Window.file_path_handing(self)
                    result = pdf.pdf_add_watermark(self.in_path.get(), output_dic_path)
                    Window.finsh_or_error(self, result)
                    time.sleep(1)
                elif self.var.get() == '4':  # docx加水印
                    file_type = main.file_type(self.in_path.get())
                    if file_type != 'docx':
                        self.finish.set("文件类型错误，仅为docx文件")
                        return
                    output_dic_path = Window.dic_path_handing(self)
                    result = word.word_add_watermark(self.in_path.get(), output_dic_path, self.watermark.get(), self.format.get(),
                                                     self.size.get())
                    Window.finsh_or_error(self, result)
                    time.sleep(1)
            else:  # 是目录
                i = 1
                for file in main.file_in_dic(self.in_path.get()):
                    if self.var.get() == '1':  # doc转docx
                        file_type = main.file_type(file)
                        if file_type != 'doc':
                            self.finish.set("文件类型错误，仅为doc文件")
                            break
                        word_object = win32com.client.Dispatch("Word.Application")
                        output_dic_path = Window.dic_path_handing(self, file)
                        result = word.doc_to_docx(file, output_dic_path, word_object)
                        word_object.Quit()
                        Window.finsh_or_error(self, result)
                        time.sleep(1)
                    elif self.var.get() == '2':  # docx转pdf
                        file_type = main.file_type(file)
                        if file_type != 'docx':
                            self.finish.set("文件类型错误，仅为docx文件")
                            break
                        word_object = win32com.client.Dispatch("Word.Application")
                        output_dic_path = Window.dic_path_handing(self, file)
                        result = word.docx_to_pdf(file, output_dic_path, word_object)
                        word_object.Quit()
                        Window.finsh_or_error(self, result)
                        time.sleep(1)
                    elif self.var.get() == '3':  # pdf加水印
                        file_type = main.file_type(file)
                        if file_type != 'pdf':
                            self.finish.set("文件类型错误，仅为pdf文件")
                            return
                        if i == 1:  # 黄健水印模板
                            pdf.creat_watermark(self.watermark.get(), self.format.get(), self.size.get())
                            i = i + 1
                        output_dic_path = Window.dic_path_handing(self, file)
                        result = pdf.pdf_add_watermark(file, output_dic_path)
                        Window.finsh_or_error(self, result)
                        time.sleep(1)
                    elif self.var.get() == '4':  # word加水印
                        file_type = main.file_type(file)
                        if file_type != 'docx':
                            self.finish.set("文件类型错误，仅为docx文件")
                            return
                        output_dic_path = Window.dic_path_handing(self, file)
                        result = word.word_add_watermark(file, output_dic_path, self.watermark.get(), self.format.get(), self.size.get())
                        Window.finsh_or_error(self, result)
                        time.sleep(1)


