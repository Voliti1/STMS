import sys
import os
import threading
from flask import Flask, request, jsonify
import pandas as pd
import openpyxl
import webview

app = Flask(__name__)

# 재고 DB 파일 이름
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

@app.route('/inventory')
def get_inventory():
    if not os.path.exists(DB_FILE):
        return jsonify([])
    try:
        df = pd.read_excel(DB_FILE, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # 출력 순서 고정
        target_order = ['순번', '신품번', '구품번', '품명', '재고수량', '공정 진행중', '입고수량']
        existing_cols = [col for col in target_order if col in df.columns]
        df = df[existing_cols]

        return jsonify(df.fillna('').to_dict(orient='records'))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/update_inventory', methods=['POST'])
def update_inventory():
    data = request.json
    if not data:
        return jsonify({"error": "전송된 데이터가 없습니다."}), 400

    try:
        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        
        # [수정됨] ws.max_column() -> ws.max_column 으로 괄호 제거
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i 
                   for i in range(1, ws.max_column + 1) if ws.cell(row=header_row_num, column=i).value}

        for row_data in data:
            p_no = str(row_data.get('신품번', '')).strip()
            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                if str(ws.cell(row=r_idx, column=col_map['신품번']).value).strip() == p_no:
                    for col_name, value in row_data.items():
                        if col_name in col_map:
                            ws.cell(row=r_idx, column=col_map[col_name]).value = value
                    break

        wb.save(DB_FILE)
        return jsonify({"message": "수정사항이 성공적으로 저장되었습니다.", "status": "success"})
    except Exception as e:
        return jsonify({"error": f"저장 실패: {str(e)}"}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files: return jsonify({"error": "파일이 없습니다."}), 400
    file = request.files['file']
    try:
        inventory_df = pd.read_excel(DB_FILE, header=2)
        inventory_df.columns = [str(c).strip() for c in inventory_df.columns]
        delivery_df = pd.read_excel(file)
        delivery_df.columns = [str(c).strip() for c in delivery_df.columns]
        delivery_df['품번'] = delivery_df['품번'].astype(str).str.strip()
        delivery_df['납품수량'] = pd.to_numeric(delivery_df['납품수량'], errors='coerce').fillna(0)
        summary = delivery_df.groupby('품번')['납품수량'].sum().reset_index()

        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
        p_no_col, stock_col = col_map.get('신품번'), col_map.get('재고수량')

        updated_count = 0
        for _, row in summary.iterrows():
            p_no, qty = row['품번'], row['납품수량']
            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                if str(ws.cell(row=r_idx, column=p_no_col).value).strip() == p_no:
                    raw_val = ws.cell(row=r_idx, column=stock_col).value
                    try:
                        cur_qty = float(str(raw_val).replace(',', '')) if raw_val else 0
                        ws.cell(row=r_idx, column=stock_col).value = cur_qty - qty
                        updated_count += 1
                    except: pass
                    break
        wb.save(DB_FILE)
        return jsonify({"message": f"{updated_count}개 품목 재고 차감 완료", "status": "success"})
    except Exception as e: return jsonify({"error": str(e)}), 500

def run_flask():
    app.run(host='127.0.0.1', port=5000)

if __name__ == '__main__':
    threading.Thread(target=run_flask, daemon=True).start()
    webview.create_window('STMS 재고 관리 시스템', 'http://127.0.0.1:5000', width=1200, height=850)
    webview.start()