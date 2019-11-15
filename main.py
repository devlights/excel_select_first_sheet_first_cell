#################################################################
# 指定されたフォルダ配下のExcelを開いていき最初のシートの最初のセルを選択状態にします.
#
# 実行には、以下のライブラリが必要です.
#   - win32com
#     - $ python -m pip install pywin32
#
# [参考にした情報]
#################################################################
import argparse


# noinspection SpellCheckingInspection
def go(target_dir: str):
    import pathlib

    import pywintypes
    import win32com.client

    excel_dir = pathlib.Path(target_dir)
    if not excel_dir.exists():
        print(f'target directory not found [{target_dir}]')
        return

    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True

        for f in excel_dir.glob('**/*.xlsx'):
            abs_path = str(f)
            try:
                wb = excel.Workbooks.Open(abs_path)
                wb.Activate()
            except pywintypes.com_error as err:
                print(err)
                continue

            try:
                ws = wb.Worksheets(1)
                ws.Activate()
                ws.Cells.Item(1, 1).Select()
                wb.Save()
                wb.Saved = True
            finally:
                wb.Close()
    finally:
        excel.Quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        usage='python main.py -d /path/to/excel/dir',
        description='Excelを開いて最初のシートの最初のセルを選択状態にします。',
        add_help=True
    )

    parser.add_argument('-d', '--directory', help='対象ディレクトリ', required=True)

    args = parser.parse_args()

    go(args.directory)
