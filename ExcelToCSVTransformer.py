import os
import sys
import logging
import win32com.client as win32

def setup_logging():
    """ログ設定を初期化します。"""
    logging.basicConfig(filename='ExcelToCSVTransformer.log', level=logging.DEBUG,
                        format='%(asctime)s: %(levelname)s: %(message)s')

def run_macro(excel_file_path, macro_name):
    """指定されたExcelファイルでマクロを実行します。"""
    excel = None
    try:
        # Excelアプリケーションを開始
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True  # Excelのウィンドウを表示する
        logging.info("Excelアプリケーションが起動しました。")

        # マクロが含まれるワークブックを開く
        logging.info(f"ワークブックを開いています: {excel_file_path}")
        workbook = excel.Workbooks.Open(Filename=excel_file_path)

        # マクロを実行
        logging.info(f"マクロを実行しています: {macro_name}")
        excel.Application.Run(macro_name)

        # ワークブックを保存して閉じる
        workbook.Save()
        workbook.Close(True)
        logging.info("ワークブックが保存され、閉じられました。")

    except Exception as e:
        logging.exception("Excelの操作中にエラーが発生しました。")
        raise  # エラーを再発生させます

    finally:
        if excel:
            # Excelアプリケーションを閉じる
            excel.Quit()
            logging.info("Excelアプリケーションが閉じられました。")

def main():
    setup_logging()  # ログ設定を初期化

    # 実行可能ファイルのディレクトリを取得
    if getattr(sys, 'frozen', False):
        dir_path = os.path.dirname(sys.executable)
    else:
        dir_path = os.path.dirname(os.path.realpath(__file__))

    excel_file_path = os.path.join(dir_path, "ExcelToCSVTransformer.xlsm")
    macro_name = "ExcelToCSVTransformer.xlsm!ExcelToCSVTransformer"

    try:
        run_macro(excel_file_path, macro_name)
        logging.info("マクロの実行が完了しました。")
    except Exception as e:
        logging.error(f"マクロの実行に失敗しました: {e}")

if __name__ == "__main__":
    main()
