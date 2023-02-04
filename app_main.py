import argparse
from pydoc import describe
import sys
from ttvduw import DocuPrinter
from ttvduw_gui import TtvduwGui
from ttvduw_utils import select_datafeeder

def test():
    #########
    ## testing DataFeeder
    # from test_ttvduw import test_XlsxDataFeeder
    # test_XlsxDataFeeder()
    #########
    ## testing DocuPrinter
    # from test_ttvduw import test_DocuPrinter
    # test_DocuPrinter()

    #########
    # from test_ttvduw import test_all_base
    # test_all_base()

    #########
    # from test_ttvduw import test_CsvDataFeeder
    # test_CsvDataFeeder()  

    #########
    # from test_ttvduw import test_gui
    # test_gui()

    #########
    # from test_ttvduw import test_mem_stress
    # test_mem_stress(loops=50)

    pass

def list2dict(lst: list):
    if len(lst) == 0 or (len(lst) % 2) != 0 :
        raise ValueError(f'Length ({len(lst)}) of input lst {lst} is not correct.')
    d = {}
    for i in range(0, len(lst), 2):
        d[lst[i]] = lst[i+1]
    return d

def main():
    desc = '根据模板和数据批量生成文档，这就是你想要的文档。 Producing documents with given template and data. This is The Very Document yoU Want (TTVDUW) '
    parser = argparse.ArgumentParser(description=desc)
    parser.add_argument('-t', '--template', required=True, help='模板文件。docx格式')
    parser.add_argument('-f', '--data-feeder-file', required=True, help='键值数据表文件。目前支持xlsx, csv')
    parser.add_argument('-o', '--out-path', help='输出目录。如果不提供则根据 -t 指定的模板文件名生成')
    parser.add_argument('-D', '--define-key-val', nargs='+', help='定义常数键值对。参数个数必须是2或2以上的偶数')
    parser.add_argument('--tab-start-from-row', type=int, default=1, help='键值数据表文件从第几行开始有数据(default: 1)')
    parser.add_argument('--tab-start-from-col', type=int, default=1, help='键值数据表文件从第几列开始有数据(default: 1)')
    parser.add_argument('--custom-out-names-with-keys', nargs='+', help='使用哪些键的值作为输出文件名')
    
    if len(sys.argv) > 1:
        # command line mode
        args = parser.parse_args()
        the_doc = DocuPrinter(args.template, out_path=args.out_path)
        DF = select_datafeeder(args.data_feeder_file)
        with DF(
            args.data_feeder_file,
            tab_start_from_row=args.tab_start_from_row,
            tab_start_from_col=args.tab_start_from_col,
            ) as df:
            if args.define_key_val is not None:
                additional_key_val = list2dict(args.define_key_val)
            for c in df.context_feed(const_key_val=additional_key_val):
                the_doc.set_context(c)
                the_doc.write(keys=args.custom_out_names_with_keys)
    else:
        # GUI mode
        ttvduw_app = TtvduwGui()
        ttvduw_app.mainloop()


if __name__ == '__main__':
    # test()
    main()
    print('Done')