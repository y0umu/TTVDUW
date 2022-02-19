'''
Common utility functions share between modules
'''
from pathlib import PurePath
from ttvduw import XlsxDataFeeder, CsvDataFeeder

# def get_fileext(fname: str):
#     '''
#     返回fname的后缀名（小写）
#     '''
#     return PurePath(fname).suffix.lower()

def select_datafeeder(fname: str):
    '''
    根据 datafeeder 的文件名选择对应的DataFeeder子类
    '''
    ext = PurePath(fname).suffix.lower()
    ext = ext.split(sep='.')[-1]
    if ext == 'xlsx':
        DF = XlsxDataFeeder
    elif ext == 'csv':
        DF = CsvDataFeeder
    else:
        raise NotImplementedError(f'The type "{ext}" is not supported yet')
    return DF
