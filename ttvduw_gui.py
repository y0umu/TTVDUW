'''
Graphical user interface for ttvduw
'''
from os import close
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as msgbox

from docx.opc.exceptions import PackageNotFoundError
from openpyxl.utils.exceptions import InvalidFileException

from ttvduw import DocuPrinter, DataFeeder

class TtvduwGui(tk.Tk):
    tpl_filetypes = (
        ('Word 文档', '*.docx'),
    )
    df_filetypes = (
        ('Excel 工作簿', '*.xlsx'),
    )

    def __init__(self):
        super().__init__()
        self.geometry('300x400')
        self.title('这就是你想要的文档')

        self.gui_help = '''1. 选择模板的路径
2. 选择键值数据表（xlsx文件）的路径
3. 选择输出文件夹 [非必选]
4. 配置输出文件名 [非必选]
5. 点击"生成我需要的文件！"按钮'''

        self.txt_tpl = tk.StringVar()  # 所选择模板的路径
        self.txt_df = tk.StringVar()   # 所选择键值数据表DataFeeder文件（xlsx等）文件的路径
        self.txt_outdir = tk.StringVar() # 输出目录
        self.txt_generate = tk.StringVar() # “生成”按钮显示的文本。可能值：['生成我需要的文档！', '生成中...']
        self.txt_generate.set('生成我需要的文档！')
        
        # grid config
        self.columnconfigure(0, weight=1)
        # self.columnconfigure(1, weight=3)  # 下标为1的列是下表为0的列的3倍宽度

        # 模板的路径、键值数据表DataFeeder文件路径是否准备好的flag
        self.is_tpl_ready = False
        self.is_df_ready = False

        # 从表的第几行/列开始取数据（默认为1）
        self.txt_tab_start_from_row = tk.StringVar()
        self.txt_tab_start_from_row.set('1')
        self.txt_tab_start_from_col = tk.StringVar()
        self.txt_tab_start_from_col.set('1')

        # 绘制图形界面
        self.create_widgets()

    def create_widgets(self):
        # 模板选择
        self.lf_pick_tpl = ttk.LabelFrame(self, text='模板')
        # self.lf_pick_tpl.grid(column=0, row=0)
        self.lf_pick_tpl.pack(fill='x', pady=5)
        ## 选择按钮
        self.btn_filepick_tpl = ttk.Button(
            self.lf_pick_tpl, 
            text='选择文件...',
            command=self.filepick_tpl_callback
        )
        self.btn_filepick_tpl.grid(column=0, row=0)
        ## 路径显示
        self.textbox_tpl = ttk.Entry(self.lf_pick_tpl, textvariable=self.txt_tpl)
        self.textbox_tpl['state'] = 'disabled'
        self.textbox_tpl.grid(column=1, row=0)
        # self.textbox_tpl.insert(0, '请选择模板文件')

        # DataFeeder文件选择
        self.lf_pick_df = ttk.LabelFrame(self, text='键值数据表')
        # self.lf_pick_df.grid(column=0, row=1)
        self.lf_pick_df.pack(fill='x', pady=5)
        ## 选择按钮
        self.btn_filepick_df = ttk.Button(
            self.lf_pick_df, 
            text='选择文件...',
            command=self.filepick_df_callback
        )
        self.btn_filepick_df.grid(column=0, row=0)
        ## 路径显示
        self.textbox_df = ttk.Entry(self.lf_pick_df, textvariable=self.txt_df)
        self.textbox_df['state'] = 'disabled'
        self.textbox_df.grid(column=1, row=0)
        # self.textbox_df.insert(0, '请选择数据文件')
        ## 数据区域指定
        ### 行
        self.label_tab_start_from_row = ttk.Label(self.lf_pick_df, text='从表格第?行开始取数据')
        self.label_tab_start_from_row.grid(column=0, row=1)
        self.textbox_tab_start_from_row = ttk.Entry(
            self.lf_pick_df,
            textvariable=self.txt_tab_start_from_row
        )
        self.textbox_tab_start_from_row.bind('<FocusOut>', self._isnum)
        self.textbox_tab_start_from_row.grid(column=1, row=1)
        ### 列
        self.label_tab_start_from_col = ttk.Label(self.lf_pick_df, text='从表格第?列开始取数据')
        self.label_tab_start_from_col.grid(column=0, row=2)
        self.textbox_tab_start_from_col = ttk.Entry(
            self.lf_pick_df,
            textvariable=self.txt_tab_start_from_col
        )
        self.textbox_tab_start_from_col.bind('<FocusOut>', self._isnum)
        self.textbox_tab_start_from_col.grid(column=1, row=2)

        # 选择输出路径
        self.lf_outdir = ttk.LabelFrame(self, text='输出文件夹配置（可选）')
        # self.lf_outdir.grid(column=0, row=2)
        self.lf_outdir.pack(fill='x', pady=5)
        ## 选择按钮
        self.btn_outdir = ttk.Button(
            self.lf_outdir,
            text='选择文件夹...',
            command=self.outdir_pick_callback
        )
        self.btn_outdir.grid(column=0, row=0)
        ## 路径显示
        self.textbox_outdir = tk.Entry(self.lf_outdir, textvariable=self.txt_outdir)
        self.textbox_outdir['state'] = 'disabled'
        self.textbox_outdir.grid(column=1, row=0)
        # self.textbox_outdir.insert(0, '选择输出目录')
        

        # 操作说明
        self.lf_help = ttk.LabelFrame(self, text='操作说明')
        # self.lf_help.grid(column=0, row=3)
        self.lf_help.pack(fill='x', pady=5)
        ## 操作说明文本
        self.text_help = tk.Text(self.lf_help, width=35, height=7)
        self.text_help.insert('1.0', self.gui_help)
        self.text_help['state'] = 'disabled'
        self.text_help.grid(column=0, row=0)

        # 自定义输出文件名 和 开始生成
        self.fm_bottom = ttk.Frame(self)
        # self.fm_bottom.grid(column=0, row=4)
        self.fm_bottom.pack()
        ## 自定义输出文件名
        self.btn_custom_outname = ttk.Button(self.fm_bottom, text='选取输出文件名...')
        self.btn_custom_outname.grid(column=0, row=0)
        ## 开始生成
        self.btn_generate = ttk.Button(
            self.fm_bottom, 
            textvariable=self.txt_generate,
            command=self.btn_generate_callback
        )
        self.btn_generate.grid(column=1, row=0)
        
    def filepick_tpl_callback(self):
        tpl_name = filedialog.askopenfilename(filetypes=TtvduwGui.tpl_filetypes)
        # 如果用户关闭，会选择一个空字符串
        print(f'"{tpl_name}" selected as tpl_name')
        if len(tpl_name) <= 0:
            self.is_tpl_ready = False
        else:
            self.is_tpl_ready = True
        self.txt_tpl.set(tpl_name)

    def filepick_df_callback(self):
        df_name = filedialog.askopenfilename(filetypes=TtvduwGui.df_filetypes)
        # 如果用户关闭，会选择一个空字符串
        print(f'"{df_name}" selected as df_name')
        if len(df_name) <= 0:
            self.is_df_ready = False
        else:
            self.is_df_ready = True
        self.txt_df.set(df_name)

    def outdir_pick_callback(self):
        outdir = filedialog.askdirectory()
        print(f'"{outdir}" selected as outdir')
        if len(outdir) == 0:
            self.txt_outdir.set('')
        else:
            self.txt_outdir.set(outdir)
    
    def btn_generate_callback(self):
        tpl_name = self.txt_tpl.get()
        df_name = self.txt_df.get()
        outdir = self.txt_outdir.get()
        row_start = int(self.txt_tab_start_from_row.get())
        col_start = int(self.txt_tab_start_from_col.get())

        try:
            the_doc = DocuPrinter(tpl_name, out_path=outdir)
            with DataFeeder(df_name,
                            tab_start_from_row=row_start,
                            tab_start_from_col=col_start,
                        ) as df:
                for c in df.context_feed():
                    the_doc.set_context(c)
                    the_doc.write()
            print('Generation of your very documents are done.')
        except PackageNotFoundError as e_docx:
            print(f'Err: {e_docx.args[0]}. Did you specify the template path correctly?')
            msgbox.showerror(title='docx文件问题', message='是否正确选取了作为模板的docx文件？')
        except InvalidFileException as e_xlsx:
            print(f'Err: {e_xlsx.args[0]}. Did you specify the xlsx file correctly?')
            msgbox.showerror(title='xlsx文件问题', message='是否正确选取了作为键值表的xlsx文件？')
        except:
            msgbox.showwarning(title='警告', message='遇到了未测试过的问题，详情请查看控制台')
            raise
        finally:
            # some cleannings here
            pass

    def _isnum(self, *args):
        '''
        表格起始行/列输入正确性坚持
        '''
        col_str = self.txt_tab_start_from_col.get()
        row_str = self.txt_tab_start_from_row.get()
        if col_str.isdigit() == False:
            self.txt_tab_start_from_col.set('1')
            msgbox.showerror(title='？？？', message='你应该输入数字')
            return
        if row_str.isdigit() == False:
           self.txt_tab_start_from_row.set('1')
           msgbox.showerror(title='？？？', message='你应该输入数字')
           return
        col = int(col_str)
        row = int(row_str)
        if col <= 0:
            self.txt_tab_start_from_col.set('1')
            msgbox.showerror(title='？？？', message='你应该输入不小于1的整数')
            return
        if row <= 0:
            self.txt_tab_start_from_row.set('1')
            msgbox.showerror(title='？？？', message='你应该输入不小于1的整数')
            return
        


