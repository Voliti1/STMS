import sys
import os
import threading
from flask import Flask, request, jsonify
import pandas as pd
import openpyxl
import webview

app = Flask(__name__)

# 재고 DB 파일 이름 (exe와 같은 폴더 위치)
DB_FILE = '부품_재고현황.xlsx'

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

@app.route('/')
def index():
    path = resource_path('index.html')
    try:
        return open(path, encoding='utf-8').read()
    except FileNotFoundError:
        return "index.html 파일을 찾을 수 없습니다."

@app.route('/inventory', methods=['GET'])
def get_inventory():
    if not os.path.exists(DB_FILE):
        return jsonify([])
    try:
        # 3행이 헤더이므로 header=2 설정
        df = pd.read_excel(DB_FILE, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        # 요청하신 컬럼 순서 고정
        target_order = ['순번', '신품번', '구품번', '품명', '재고수량', '공정 진행중', '입고수량']
        existing_cols = [c for c in target_order if c in df.columns]
        df = df[existing_cols]
        
        df = df.fillna('')
        return jsonify(df.to_dict(orient='records'))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files or not os.path.exists(DB_FILE):
        return jsonify({"error": "파일이 없거나 설정이 잘못되었습니다."}), 400

    try:
        file = request.files['file']
        delivery_df = pd.read_excel(file)
        delivery_df.columns = [str(c).strip() for c in delivery_df.columns]
        delivery_df['품번'] = delivery_df['품번'].astype(str).str.strip()
        delivery_df['납품수량'] = pd.to_numeric(delivery_df['납품수량'], errors='coerce').fillna(0)
        summary = delivery_df.groupby('품번')['납품수량'].sum().reset_index()

        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        # 'int' object is not callable 에러 해결: 괄호 제거
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i 
                   for i in range(1, ws.max_column + 1)}

        p_no_col = col_map['신품번']
        stock_col = col_map['재고수량']

        updated_count = 0
        for _, row in summary.iterrows():
            p_no, sub_qty = row['품번'], row['납품수량']
            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                if str(ws.cell(row=r_idx, column=p_no_col).value).strip() == p_no:
                    raw_val = ws.cell(row=r_idx, column=stock_col).value
                    curr_qty = float(str(raw_val).replace(',', '')) if raw_val not in [None, ''] else 0
                    ws.cell(row=r_idx, column=stock_col).value = curr_qty - sub_qty
                    updated_count += 1
                    break
        
        wb.save(DB_FILE)
        return jsonify({"message": f"{updated_count}개 품목 차감 완료", "status": "success"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/update_inventory', methods=['POST'])
def update_inventory():
    try:
        updated_data = request.json
        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i 
                   for i in range(1, ws.max_column + 1)}

        for item in updated_data:
            p_no = str(item.get('신품번', '')).strip()
            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                if str(ws.cell(row=r_idx, column=col_map['신품번']).value).strip() == p_no:
                    for col, val in item.items():
                        if col in col_map: ws.cell(row=r_idx, column=col_map[col]).value = val
                    break
        wb.save(DB_FILE)
        return jsonify({"status": "success"})
    except Exception as e: return jsonify({"error": str(e)}), 500

def run_flask(): app.run(host='127.0.0.1', port=5000)

if __name__ == '__main__':
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()
    webview.create_window('STMS 재고 관리 시스템', 'http://127.0.0.1:5000', width=1200, height=900)
    webview.start()