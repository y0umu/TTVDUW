import argparse
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
    pass

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--template', required=True, help='模板文件。docx格式')
    parser.add_argument('-f', '--data-feeder-file', required=True, help='键值数据表文件。目前支持xlsx, csv')
    parser.add_argument('-o', '--out-path', help='输出目录。如果不提供则根据 -t 指定的模板文件名生成')
    parser.add_argument('--tab-start-from-row', type=int, default=1, help='键值数据表文件从第几行开始有数据(default: 1)')
    parser.add_argument('--tab-start-from-col', type=int, default=1, help='键值数据表文件从第几列开始有数据(default: 1)')
    parser.add_argument('--custom-out-names-with-keys', nargs='+', help='使用哪些键的值作为输出文件名')
    
    if len(sys.argv) > 1:
        # command line mode
        args = parser.parse_args()
        the_doc = DocuPrinter(args.template, out_path=args.out_path)
        DF = select_datafeeder(args.data_feeder_file)
        with DF(args.data_feeder_file,
                        tab_start_from_row=args.tab_start_from_row,
                        tab_start_from_col=args.tab_start_from_col,
                    ) as df:
            for c in df.context_feed():
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