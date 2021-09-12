'''
This is The Very Document yoU Want (ttvduw)
'''
from pathlib import Path
import time
from copy import deepcopy
from docxtpl import DocxTemplate
from openpyxl import load_workbook

class DocuPrinter():
    '''
    读取 模板.docx，接受DataFeeder传入数据，输出格式化好的docx
    '''
    def __init__(self, tpl_name, out_path=None):
        '''
        tpl_name: 模板文件名。是一个docx文件
        out_path: 输出目录, str或None。如果不传入从str值则会将tpl_name的基名（PurePath.stem）作为输出目录的名称
        '''
        self.docu = DocxTemplate(tpl_name)
        self._ori_docx = deepcopy(self.docu.docx)   # hack。让同一个DocxTemplate()能多次使用
        self.context = {}  # 要填充的键值对，由self.set_context正确设置
        
        if out_path is None or len(out_path) == 0:
            out_path = Path(tpl_name).stem
        p_out_path = Path(out_path)
        p_out_path.mkdir(exist_ok=True)
        self.p_out_path = p_out_path

    
    def set_context(self, context: dict):
        '''
        context: 获取要填充的实际数据，每次获取的数据只填充一篇文档
        '''
        self.context.update(context)

    def write(self, keys=None):
        '''
        将渲染（填充）后的文件写入磁盘。写入的文件名从self.context中的键对应值选择（见keys参数）

        keys: list. 选择哪些键作为参数

        警告：这个方法不负责检查要输出的文件是否已经存在。如果已经存在，原文件将被覆盖。
        '''
        docu = self.docu
        docu.render(self.context)

        # 设置输出文件名
        out_name = ""
        use_fallback_filename = False
        if keys is None:
            use_fallback_filename = True
        else:
            for k in keys:
                _s = self.context.get(k)
                if _s is None:   # 用户输入了非法的键，跳过
                    continue
                out_name += str(_s)
                out_name += '_'
            if len(out_name) == 0: # 用户输入的全是非法键
                use_fallback_filename = True
        
        if use_fallback_filename:
            print('WARN: No valid keys given. Use fallback output filename')
            out_name = str(time.time() )   # fallback filename
        
        out_name += '.docx'
        out_name = str(self.p_out_path / Path(out_name))
        docu.save(out_name)
        self.docu.docx = deepcopy(self._ori_docx)   # 重置DocxTemplate().docx使这个模板能再次使用
    

class DataFeeder():
    '''
    读取数据表格的接口类
    '''
    def __init__(self, fname: str, ftype='xlsx', 
                 tab_start_from_row=1, tab_start_from_col=1):
        '''
        fname: str. 输入数据文件路径
        ftype: str. 目前只能选择 'xlsx'
        tab_start_from_row: int. 数据从第几行开始（默认：1）
        tab_start_from_col: int. 数据从第几列开始（默认：1）
        '''
        self.fname = fname
        self.ftype = ftype
        self.min_row = tab_start_from_row
        self.min_col = tab_start_from_col
        self.keys = None     # 存储读取到的表格区键名，self._load_with_xlsx将正确设置此变量
        # self.data_gen = None # 存储键值的生成器，self._DataGen将正确设置此变量
        self.wb_xlsx = None  # 存储xlsx类型, ftype == 'xlsx'时使用
        if ftype == 'xlsx':
            self.wb_xlsx = self._load_with_xlsx()
        else:
            raise NotImplementedError("Support of file type {} is not yet implemented".format(ftype))
    
    def __enter__(self):
        return self
    
    def __exit__(self, *exc_args):
        '''
        exc_args: tuple of (exc_type, exc_val, exc_tb)
        '''
        if self.ftype == 'xlsx':
            self.wb_xlsx.close()
            print('debug: DataFeeder cleanned')

    def context_feed(self):
        '''
        这是一个生成器，喂给DocxTemplate.get_context()
        '''
        context = {}
        data_gen = self._DataGen()
        for data in data_gen:
            for k,v in zip(self.keys, data):
                context[k] = v
            yield context

    def _DataGen(self):
        if self.ftype == 'xlsx':
            keys_row, data_rows = self._get_ws_key_data_rows()
            # 将表格中的行转换成python list
            for r in data_rows:   # 余下的是每个记录具体的值
                ## 争议：空单元格应该直接返回None（默认行为）还是改写为""（空字符串）
                # 返回None时，如果模板中没有任何条件判断，就会印出"None"这四个字母
                # 返回""时，会让模板中对应位置什么有没有
                # this_row = [ x.value for x in r]
                this_row = [ x.value if x.value is not None else "" for x in r]
                this_row = this_row[self.min_col-1:]  # discarding unwanted colums
                yield this_row
        else:
            raise NotImplementedError("Support of file type {} is not yet implemented".format(ftype))
    
    def _get_ws_key_data_rows(self):
        '''
        获取xlsx工作表中的键行，以及数据行生成器
        
        return:
        keys_row, ws_data_rows
        '''
        if self.ftype == 'xlsx':
            wb = self.wb_xlsx
            ws = wb.active
            ws_data_rows = ws.iter_rows()
            # discarding the unwanted rows
            if self.min_row > 1:
                for _i in range(self.min_row - 1):
                    next(ws_data_rows)
            keys_row = next(ws_data_rows)  # 认为表格区的第1行是键名（字段名）
            return keys_row, ws_data_rows
        else:
            raise UserWarning('You should not call this method if not using xlsx')

    def _load_with_xlsx(self):
        wb = load_workbook(self.fname, read_only=True, data_only=True)
        self.wb_xlsx = wb
        keys_row, data_row = self._get_ws_key_data_rows()

        keys = [ str(x.value) for x in keys_row ]  # 转换成 python list
        ## 暂时不能跳过空单元格的原因是：
        ## self.context_feed 里面用了zip将keys与每一行一一对应
        ## 跳过keys那一行的空单元格（“空”键）会导致非空单元格（非“空”键）与其值不能对应上
        ## （“空”键与值的对应无论如何都是错的）
        # keys = []
        # for x in keys_row:
        #     if x.value is None:  # 跳过空单元格
        #         continue
        #     keys.append(str(x.value))
        self.keys = keys
        # self.data_gen = self._DataGen()
        
        return wb
    
