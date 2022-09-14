import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import sqlite3


class MyApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        ##ウィンドウサイズ
        width = 600
        height = 300
        self.geometry(f'{width}x{height}') #作成するアプリ画面(Window)の、位置や大きさを調整するために利用される関数
        self.minsize(width, height)
        self.maxsize(width, height)
        self.title(f'DnD')

        ##フレーム
        self.frame_drag_drop = frameDragAndDrop(self) # ここで関数を呼び出す
        ## 配置
        self.frame_drag_drop.grid(column=0, row=0, padx=5, pady=5, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1 )
        
class frameDragAndDrop(tk. LabelFrame):
    def __init__(self, parent):
        super().__init__(parent)
        conn = sqlite3.connect('test.db')
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS path(id INTEGER PRIMARY KEY AUTOINCREMENT, name STRING)')
        #pathを取得
        self.output_path = cur.execute("SELECT name FROM path WHERE id=1 ")
        final_path = ""
        for i in self.output_path:
            final_path = i[0]
        self.textbox = tk.Text(self, background='#DFDFFF')
        self.textbox.insert('0.0', "ここにファイルをドロップしてください\n\n")
        # self.textbox.insert('end', "出力先：" + final_path)
        self.textbox.configure(state='disabled')
        self.button1 = tk.Button(text="出力先のpathを変更", command=self.reload)
        self.button1.place(x=450, y=265)
        self.lbl = tk.Label(text='出力先：')
        self.lbl.place(x=20, y=270)
        self.txt = tk.Entry(width=60)
        self.txt.place(x=80, y=270)
        self.txt.insert(0, final_path)

        ##ドラックアンドドロップ
        self.textbox.drop_target_register(DND_FILES)
        self.textbox.dnd_bind('<<Drop>>', self.funcDragAndDrop)

        ##スクロールバー設定
        self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.textbox.yview)
        self.textbox['yscrollcommand'] = self.scrollbar.set

        ##配置
        self.textbox.grid(column=0, row=0, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def reload(self):
        output = self.txt.get()
        if output == "":
            pass
        else:
            dbname = 'test.db'
            #カレントディレクトリにtest.dbがなければ作成する、既存の場合は接続
            conn = sqlite3.connect(dbname)
            #sqliteを操作するカーソルオブジェクト
            cur = conn.cursor()
            #テーブルのCreate文を実行

            cur.execute('CREATE TABLE IF NOT EXISTS path(id INTEGER PRIMARY KEY AUTOINCREMENT, name STRING)')
            delete_sql = 'DELETE FROM path WHERE id=1'
            insert_sql = 'INSERT INTO path values(1,'+"'"+output+"'"+')'
            print(output)
            cur.execute(delete_sql)
            cur.execute(insert_sql)
            conn.commit()
            
            cur.close()
            conn.close()
            # self.output_path = str(txt)
            self.lbl = tk.Label(text='出力先：')
            self.lbl.place(x=20, y=270)
            #テキストボックス
            self.txt = tk.Entry(width=60)
            self.txt.insert(0, output)
            self.txt.place(x=80, y=270)

    def funcDragAndDrop(self, e):
        ##ここを編集
        message= e.data
        self.textbox.configure(state='normal')
        self.textbox.insert(tk.END, message)
        self.textbox.configure(state='disabled')
        path = conversion(message)
        path = path.change()
        self.textbox.see(tk.END)

class conversion():
    def __init__(self, message):
        self.file_path = message
    
    def change(self):
        conn = sqlite3.connect('test.db')
        cur = conn.cursor()
        #pathを取得
        output_path = cur.execute("SELECT name FROM path WHERE id=1 ")
        for i in output_path:
            final_path = i[0]
        df = pd.read_csv(self.file_path, header=None)
        df_lst = list(pd.read_csv(self.file_path))
        if "target_str1" in df_lst[11]:
            df.to_excel(final_path+'\target_str1.xlsx', encoding='utf-8', index=False)
        elif "target_str2" in df_lst[11]:
            df.to_excel(final_path+'\target_str2.xlsx', encoding='utf-8', index=False)
        elif "target_str3" in df_lst[12]:
            df.to_excel(final_path+'\target_str3.xlsx', encoding='utf-8', index=False)
        elif "target_str4" in df_lst[12]:
            df.to_excel(final_path+'\target_str4.xlsx', encoding='utf-8', index=False)
        elif "target_str5" in df_lst[12]:
            df.to_excel(final_path+'\target_str5.xlsx', encoding='utf-8', index=False) 
        elif "target_str6" in df_lst[11]:
            df.to_excel(final_path+'\target_str6.xlsx', encoding='utf-8', index=False)        
        elif "target_str7" in df_lst[11]:
            df.to_excel(final_path+'\target_str7.xlsx', encoding='utf-8', index=False)
        elif "target_str8" in df_lst[11]:
            df.to_excel(final_path+'\target_str8.xlsx', encoding='utf-8', index=False)
        conn.close()

if __name__ == "__main__":
    app = MyApp()
    result = app.mainloop()