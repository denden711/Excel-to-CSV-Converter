import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, Listbox, MULTIPLE, filedialog, messagebox, StringVar, Frame
from tkinterdnd2 import TkinterDnD, DND_FILES

# デフォルトの置換規則
replacements = {
    '40': '040',
    '60': '060',
    '80': '080',
    '30': '030',
    '50': '050',
    '100': '100',
    'なし': 'x2=000',
    ' ': ' '  # 空白はそのまま
}

def replace_name(original_name):
    """
    ファイル名を置換規則に従って変換する関数
    """
    try:
        for old, new in replacements.items():
            original_name = original_name.replace(old, new)
        return original_name
    except Exception as e:
        print(f"ファイル名の置換中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"ファイル名の置換中にエラーが発生しました: {e}")
        return original_name

def convert_xlsx_to_csv(file_paths, output_directory):
    """
    ExcelファイルをCSVファイルに変換する関数
    """
    for file_path in file_paths:
        try:
            xls = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"{file_path} を読み込めませんでした: {e}")
            messagebox.showerror("エラー", f"{file_path} を読み込めませんでした: {e}")
            continue
        
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        new_base_name = replace_name(base_name)
        
        try:
            # 最初のシートのみを読み込む
            df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        except Exception as e:
            print(f"{file_path} のシートを読み込めませんでした: {e}")
            messagebox.showerror("エラー", f"{file_path} のシートを読み込めませんでした: {e}")
            continue
        
        csv_filename = f"{new_base_name}.csv"
        csv_path = os.path.join(output_directory, csv_filename)
        
        try:
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')  # UTF-8エンコーディングを指定
            print(f"{os.path.basename(file_path)} を {csv_filename} に変換しました")
        except Exception as e:
            print(f"{csv_filename} を書き込めませんでした: {e}")
            messagebox.showerror("エラー", f"{csv_filename} を書き込めませんでした: {e}")

def drop(event):
    """
    ドラッグ＆ドロップでファイルをリストに追加する関数
    """
    try:
        files = root.tk.splitlist(event.data)
        for file in files:
            if file not in file_listbox.get(0, 'end'):
                file_listbox.insert('end', file)
    except Exception as e:
        print(f"ファイルの追加中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"ファイルの追加中にエラーが発生しました: {e}")

def start_conversion():
    """
    ファイルドロップ完了ボタンを押したときに変換を開始する関数
    """
    try:
        files = file_listbox.get(0, 'end')
        if not files:
            messagebox.showinfo("情報", "ファイルが選択されていません")
            return
        
        output_directory = filedialog.askdirectory()
        if not output_directory:
            messagebox.showinfo("情報", "出力ディレクトリが選択されていません")
            return
        
        convert_xlsx_to_csv(files, output_directory)
        messagebox.showinfo("情報", "変換が完了しました")
    except Exception as e:
        print(f"変換中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"変換中にエラーが発生しました: {e}")

def add_replacement():
    """
    置換要件を追加する関数
    """
    try:
        old = old_var.get()
        new = new_var.get()
        if old and new:
            replacements[old] = new
            update_replacement_listbox()
            old_var.set('')
            new_var.set('')
        else:
            messagebox.showwarning("警告", "置換要件を正しく入力してください")
    except Exception as e:
        print(f"置換要件の追加中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"置換要件の追加中にエラーが発生しました: {e}")

def remove_replacement():
    """
    選択された置換要件を削除する関数
    """
    try:
        selected_indices = replacement_listbox.curselection()
        for index in selected_indices[::-1]:  # 後ろから削除することでインデックスのシフトを防ぐ
            key = replacement_listbox.get(index).split(' -> ')[0]
            if key in replacements:
                del replacements[key]
        update_replacement_listbox()
    except Exception as e:
        print(f"置換要件の削除中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"置換要件の削除中にエラーが発生しました: {e}")

def update_replacement_listbox():
    """
    置換要件リストボックスを更新する関数
    """
    try:
        replacement_listbox.delete(0, 'end')
        for old, new in replacements.items():
            replacement_listbox.insert('end', f"{old} -> {new}")
    except Exception as e:
        print(f"置換要件リストの更新中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"置換要件リストの更新中にエラーが発生しました: {e}")

def remove_selected_files():
    """
    選択されたファイルをリストから削除する関数
    """
    try:
        selected_indices = file_listbox.curselection()
        for index in selected_indices[::-1]:  # 後ろから削除することでインデックスのシフトを防ぐ
            file_listbox.delete(index)
    except Exception as e:
        print(f"選択されたファイルの削除中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"選択されたファイルの削除中にエラーが発生しました: {e}")

def reset_file_list():
    """
    ファイルリストをリセットする関数
    """
    try:
        file_listbox.delete(0, 'end')
    except Exception as e:
        print(f"ファイルリストのリセット中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"ファイルリストのリセット中にエラーが発生しました: {e}")

root = TkinterDnD.Tk()  # DnDをサポートするTkinterウィンドウを作成
root.title("Excel to CSV Converter")

# ドロップ領域ラベル
label = Label(root, text="ここにExcelファイルをドロップしてください", width=40, height=10, bg="lightgray")
label.pack(padx=10, pady=10)

# ファイルリストボックス
file_listbox = Listbox(root, selectmode=MULTIPLE, width=80, height=10)
file_listbox.pack(padx=10, pady=10)

# 置換要件入力フォーム
frame = Frame(root)
frame.pack(padx=10, pady=10)

old_var = StringVar()
new_var = StringVar()

old_label = Label(frame, text="置換前")
old_label.grid(row=0, column=0, padx=5, pady=5)
old_entry = Entry(frame, textvariable=old_var)
old_entry.grid(row=0, column=1, padx=5, pady=5)

new_label = Label(frame, text="置換後")
new_label.grid(row=0, column=2, padx=5, pady=5)
new_entry = Entry(frame, textvariable=new_var)
new_entry.grid(row=0, column=3, padx=5, pady=5)

add_button = Button(frame, text="追加", command=add_replacement)
add_button.grid(row=0, column=4, padx=5, pady=5)

# 置換要件リストボックス
replacement_listbox = Listbox(root, selectmode=MULTIPLE, width=80, height=10)
replacement_listbox.pack(padx=10, pady=10)

remove_button = Button(root, text="選択した置換要件を削除", command=remove_replacement)
remove_button.pack(padx=10, pady=10)

# ファイルリスト操作ボタン
file_button_frame = Frame(root)
file_button_frame.pack(padx=10, pady=10)

remove_file_button = Button(file_button_frame, text="選択したファイルを削除", command=remove_selected_files)
remove_file_button.grid(row=0, column=0, padx=5, pady=5)

reset_file_button = Button(file_button_frame, text="ファイルリストをリセット", command=reset_file_list)
reset_file_button.grid(row=0, column=1, padx=5, pady=5)

# ファイルドロップ完了ボタン
convert_button = Button(root, text="ファイルドロップ完了", command=start_conversion)
convert_button.pack(padx=10, pady=10)

# ドロップ領域の設定
label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', drop)

# 初期の置換要件リストボックスの更新
update_replacement_listbox()

root.mainloop()

