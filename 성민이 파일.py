import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt

class InventoryManager:
    def __init__(self, root):
        self.root = root
        self.root.title('부품 재고 관리 시스템')
        self.root.geometry('1200x700')
        self.manual_mode = False
        self.inventory_file = '부품_재고현황.csv'
        self.history_file = 'inventory_history.csv'
        self.load_inventory()
        self.build_ui()

    def load_inventory(self):
        self.raw = pd.read_csv(self.inventory_file)
        self.inventory = self.raw.iloc[2:, [1,3,4]].copy()
        self.inventory.columns = ['품번','품명','재고수량']
        self.inventory['재고수량'] = pd.to_numeric(self.inventory['재고수량'], errors='coerce').fillna(0)

    def build_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(fill='both', expand=True)

        tk.Button(frame, text='납품명세서 업로드', command=self.upload_delivery).pack(pady=5)
        tk.Button(frame, text='수동모드', command=self.toggle_manual).pack(pady=5)
        tk.Button(frame, text='대시보드', command=self.dashboard).pack(pady=5)

        self.tree = ttk.Treeview(frame, columns=('품번','품명','재고수량'), show='headings')
        for col in ('품번','품명','재고수량'):
            self.tree.heading(col, text=col)
        self.tree.pack(fill='both', expand=True)
        self.refresh_tree()

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in self.inventory.iterrows():
            self.tree.insert('', 'end', values=list(row))

    def upload_delivery(self):
        file = filedialog.askopenfilename(filetypes=[('CSV files','*.csv')])
        if not file:
            return
        delivery = pd.read_csv(file)
        for _, row in delivery.iterrows():
            part = row['품번']
            qty = row['납품수량']
            idx = self.inventory[self.inventory['품번']==part].index
            if len(idx):
                self.inventory.loc[idx,'재고수량'] -= qty
        self.save_inventory()
        self.refresh_tree()
        messagebox.showinfo('완료','재고 업데이트 완료')

    def save_inventory(self):
        date = datetime.today().strftime('%Y-%m-%d')
        self.raw.columns.values[6] = date
        self.raw.iloc[2:,4] = self.inventory['재고수량'].values
        self.raw.to_csv(self.inventory_file,index=False)
        log = self.inventory[['품번','재고수량']].copy()
        log['날짜']=date
        if os.path.exists(self.history_file):
            old = pd.read_csv(self.history_file)
            log = pd.concat([old,log])
        log.to_csv(self.history_file,index=False)

    def toggle_manual(self):
        self.manual_mode = not self.manual_mode
        messagebox.showwarning('수동모드', '수동 수정 가능' if self.manual_mode else '수동 수정 잠금')

    def dashboard(self):
        if not os.path.exists(self.history_file):
            messagebox.showerror('오류','이력 없음')
            return
        hist = pd.read_csv(self.history_file)
        summary = hist.groupby('날짜')['재고수량'].sum()
        summary.plot(kind='line')
        plt.title('재고 추이')
        plt.show()

root = tk.Tk()
app = InventoryManager(root)
root.mainloop()