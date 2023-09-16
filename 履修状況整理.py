import pandas as pd
import tkinter as tk
import tkinter.messagebox as messagebox

class CsvError(Exception):
    pass

def is_error(df):
    if df.shape[1] != 15:
        raise CsvError

def export_to_excel(df):
    excel_writer = pd.ExcelWriter('excel/単位取得状況.xlsx', engine='xlsxwriter')
    df.to_excel(excel_writer, sheet_name='各開講科目ごとの単位取得状況', index=False, encoding="shift-jis")

    df2 = df[df['合否'] != '否'][['科目詳細区分', '科目小区分', '単位数', '合否']]
    df2 = df2.groupby(['科目詳細区分', '科目小区分'])['単位数'].sum().reset_index()
    df2.to_excel(excel_writer, sheet_name='各科目区分の単位取得状況', index=False)

    excel_writer.save()

def semester_order(semester):
    order = {'春学期': 0, '夏学期': 1, '秋学期': 2, '冬学期': 3}
    return order.get(semester, 4)

def data_process():
    try:
        import_file_path = "csv/単位取得状況.csv"
        df = pd.read_csv(import_file_path, index_col=False, encoding="shift-jis")
        is_error(df)

        # 学期のカスタム順序でソート
        df['ソートキー'] = df['修得学期'].map(semester_order)
        df_sorted = df.sort_values(by=['修得年度', 'ソートキー'])
        del df_sorted['ソートキー']

        export_to_excel(df_sorted)
        
        tk.Tk().withdraw()
        messagebox.showinfo('メッセージ', 'data exported to "excel/単位取得状況.xlsx"')

    except FileNotFoundError:
        tk.Tk().withdraw()
        messagebox.showinfo('メッセージ', 'ファイルが指定された場所にありません')
    except CsvError:
        tk.Tk().withdraw()
        messagebox.showinfo('メッセージ', 'CSVの形式が正しくありません')

if __name__ == '__main__':
    data_process()
