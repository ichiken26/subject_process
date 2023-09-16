import pandas as pd
import tkinter as tk
import tkinter.messagebox as messagebox

class csvError(Exception):
    pass

class rowlostError(Exception):
    pass

class passfailError(Exception):
    pass

class colmissingError(Exception):
    pass

def csvchecker(df):
    required_columns = ['科目詳細区分', '科目小区分', '修得年度', '修得学期', '開講科目名', '単位数', '評語', '合否']
    missing_columns = [col for col in required_columns if col not in df.columns]
    invalid_values = df[~df['合否'].isin(['合', '否'])]
    if df.shape[1] != 15:
        raise csvError
    if missing_columns:
        return missing_columns
        raise rowlostError
    if not invalid_values.empty:
        raise passfailError
    missing_col = []
    for col in required_columns:
        if df[col].isnull().any():
            missing_col.append(df[col])
    if missing_col:
        return missing_col
        raise colmissingError

def export_to_excel(df):
    excel_writer = pd.ExcelWriter('excel/単位取得状況.xlsx', engine='xlsxwriter')
    df.to_excel(excel_writer, sheet_name='各開講科目ごとの単位取得状況', index=False)

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
        df = pd.read_csv(import_file_path, index_col=False, encoding="utf-8")
        csvchecker(df)

        # 学期のカスタム順序でソート
        df['ソートキー'] = df['修得学期'].map(semester_order)
        df_sorted = df.sort_values(by=['修得年度', 'ソートキー'])
        del df_sorted['ソートキー']

        export_to_excel(df_sorted)

    except FileNotFoundError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', 'ファイルが指定された場所にありません')
    except csvError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', 'CSVの形式が正しくありません')
    except UnicodeDecodeError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', 'エンコードが正しくありません。"UTF-8"で保存してください')
    except rowlostError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', f"必要な列が不足しています: {', '.join(csvchecker(df))}")
    except passfailError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', '合否列に "合" または "否" 以外の値が含まれています。')
    except pd.errors.ParserError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', 'CSVが正しく編集されていません')
    except colmissingError:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', f'カラム "{csvchecker(df)}"に空の情報があります')
    except:
        tk.Tk().withdraw()
        messagebox.showerror('エラー', f'予期せぬエラーが発生しました。開発者にお問い合わせください')
    else:
        tk.Tk().withdraw()
        messagebox.showinfo('メッセージ', 'data exported to "excel/単位取得状況.xlsx"')
        
        

if __name__ == '__main__':
    data_process()
