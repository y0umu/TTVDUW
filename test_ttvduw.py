from ttvduw import DocuPrinter, XlsxDataFeeder, CsvDataFeeder
from ttvduw_gui import TtvduwGui

def test_DocuPrinter():
    context = {
        'stu_name': '这是姓名',
        'stu_id': 2210605101,
        'stu_colledge': '力学与建筑工程学院',
        'stu_grade': '2022',
        'terms_done': '6',
        'weighted_avg_mark': 87.06,
        'stu_major': '智能建造',
        'rank_in_major': 1,
        'total_in_major': '20'
    }
    the_doc = DocuPrinter('examples/成绩排名证明/成绩排名证明（推免）模板_tpl.docx')
    the_doc.set_context(context)
    the_doc.write(['stu_name'])

    context.update(
        {
            'stu_name': '这也是姓名',
            'stu_id': 2210605102,
            'weighted_avg_mark': 88.00,
            'rank_in_major': 0
        }
    )
    the_doc.set_context(context)
    the_doc.write(['stu_name'])

def test_XlsxDataFeeder():
    with XlsxDataFeeder('examples/成绩排名证明/2022级智能建造学生成绩排名_datafeed.xlsx', 
                    tab_start_from_row=2) as df:
        for c in df.context_feed():
            print(c)

def test_CsvDataFeeder():
    with CsvDataFeeder('examples/成绩排名证明/2022级智能建造学生成绩排名_datafeed.csv', 
                    tab_start_from_row=2) as df:
        for c in df.context_feed():
            print(c)

def test_all_base():
    the_doc = DocuPrinter('examples/成绩排名证明/成绩排名证明（推免）模板_tpl.docx')
    with XlsxDataFeeder('examples/成绩排名证明/2022级智能建造学生成绩排名_datafeed.xlsx', 
                    tab_start_from_row=2) as df:
        for c in df.context_feed():
            the_doc.set_context(c)
            the_doc.write(keys=('这个键不存在'))
            # the_doc.write(keys=('stu_id', 'stu_name', '这个键不存在'))


def test_gui():
    ttvduw_app = TtvduwGui()
    ttvduw_app.mainloop()

def test_mem_stress(loops=1):
    '''
    This test function simulates what is described in 
    https://github.com/y0umu/TTVDUW/issues/1
    '''
    from ttvduw_dbg import total_size
    import pdb
    import gc

    # gc.set_debug(gc.DEBUG_LEAK)
    tpl_name = 'examples/成绩排名证明/成绩排名证明（推免）模板_tpl.docx'
    i = 0
    while i < loops:
        docu_printer = DocuPrinter(tpl_name)
        df = CsvDataFeeder('examples/成绩排名证明/2022级智能建造学生成绩排名_datafeed.csv', 
                        tab_start_from_row=2)
        for c in df.context_feed():
            docu_printer.set_context(c)
            docu_printer.write(keys=('stu_id', 'stu_name'))
        # print(f'bytes of docu_printer after iter {i} == {total_size(docu_printer)}')
        # print(f'bytes of df after iter {i} == {total_size(df)}')
        unreachable_cnts = gc.collect()
        print(f'end of iter {i}, unreachable_cnts was {unreachable_cnts}')
        # pdb.set_trace()
        i += 1
    
