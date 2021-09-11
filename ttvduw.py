'''
This is The Very Document yoU Want (ttvduw)
'''
import argparse
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
        
        if out_path is None:
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
        self.data_gen = None # 存储键值的生成器，self._DataGen将正确设置此变量
        self.wb_xlsx = None  # 存储xlsx类型, ftype == 'xlsx'时使用
        if ftype == 'xlsx':
            self.wb_xlsx = self._load_with_xlsx()
        else:
            raise NotImplementedError("Support of file type {} is not yet implemented".format(ftype))
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.ftype == 'xlsx':
            self.wb_xlsx.close()
            # print('debug: DataFeeder cleanned')

    def context_feed(self):
        '''
        这是一个生成器，喂给DocxTemplate.get_context()
        '''
        context = {}
        for data in self.data_gen:
            for k,v in zip(self.keys, data):
                context[k] = v
            yield context

    def _DataGen(self):
        if self.ftype == 'xlsx':
            # 将表格中的行转换成python list
            for r in self._ws_rows:   # 余下的是每个记录具体的值
                ## 争议：空单元格应该直接返回None（默认行为）还是改写为""（空字符串）
                # 返回None时，如果模板中没有任何条件判断，就会印出"None"这四个字母
                # 返回""时，会让模板中对应位置什么有没有
                # this_row = [ x.value for x in r]
                this_row = [ x.value if x.value is not None else "" for x in r]
                this_row = this_row[self.min_col-1:]  # discarding unwanted colums
                yield this_row
        else:
            raise NotImplementedError("Support of file type {} is not yet implemented".format(ftype))
        
    def _load_with_xlsx(self):
        wb = load_workbook(self.fname, read_only=True, data_only=True)
        ws = wb.active
        ws_rows = ws.iter_rows()
        # discarding the unwanted rows
        if self.min_row > 1:
            for _i in range(self.min_row - 1):
                next(ws_rows)
        keys_row = next(ws_rows)  # 认为表格区的第1行是键名（字段名）
        self._ws_rows = ws_rows

        keys = [ str(x.value) for x in keys_row ]  # 转换成 python list
        self.keys = keys
        self.data_gen = self._DataGen()
        
        return wb
    

def test():
    #########
    ## testing DataFeeder
    # from test_ttvduw import test_DataFeeder
    # test_DataFeeder()
    #########
    ## testing DocuPrinter
    # from test_ttvduw import test_DocuPrinter
    # test_DocuPrinter()

    #########
    from test_ttvduw import test_all
    test_all()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--template', required=True, help='模板文件。docx格式')
    parser.add_argument('-f', '--data-feeder-file', required=True, help='键值数据表文件。目前只支持xlsx')
    parser.add_argument('-o', '--out-path', help='输出目录。如果不提供则根据 -t 指定的模板文件名生成')
    parser.add_argument('--tab-start-from-row', type=int, default=1, help='键值数据表文件从第几行开始有数据(default: 1)')
    parser.add_argument('--tab-start-from-col', type=int, default=1, help='键值数据表文件从第几列开始有数据(default: 1)')
    parser.add_argument('--custom-out-names-with-keys', nargs='+', help='使用哪些键的值作为输出文件名')
    
    args = parser.parse_args()

    the_doc = DocuPrinter(args.template, out_path=args.out_path)
    with DataFeeder(args.data_feeder_file,
                    tab_start_from_row=args.tab_start_from_row,
                    tab_start_from_col=args.tab_start_from_col,
                   ) as df:
        for c in df.context_feed():
            the_doc.set_context(c)
            the_doc.write(keys=args.custom_out_names_with_keys)


if __name__ == '__main__':
    # test()
    main()
    print('Done')