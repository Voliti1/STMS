import sys
import os
import threading
from datetime import datetime
from flask import Flask, request, jsonify
import pandas as pd
import openpyxl
import webview

app = Flask(__name__)

# 재고 DB 파일 이름 및 폴더 설정 [cite: 370]
DB_FILE = '부품_재고현황.xlsx'
SAVE_FOLDER = '납품 명세서'

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 페이지 라우팅 ---

@app.route('/')
def landing():
    """초기 선택 페이지"""
    path = resource_path('landing.html')
    try:
        return open(path, encoding='utf-8').read()
    except FileNotFoundError:
        return "landing.html 파일을 찾을 수 없습니다."

@app.route('/inventory_system')
def inventory_system():
    """기존 재고 관리 페이지"""
    path = resource_path('index.html')
    try:
        return open(path, encoding='utf-8').read()
    except FileNotFoundError:
        return "index.html 파일을 찾을 수 없습니다."

@app.route('/dashboard')
def dashboard_page():
    """대시보드 페이지 (플레이스홀더)"""
    path = resource_path('dashboard.html')
    try:
        return open(path, encoding='utf-8').read()
    except FileNotFoundError:
        return "dashboard.html 파일을 찾을 수 없습니다."

# --- 기존 데이터 처리 로직 ---

@app.route('/inventory', methods=['GET'])
def get_inventory():
    if not os.path.exists(DB_FILE):
        return jsonify([])
    try:
        # 3행 헤더 설정 [cite: 246]
        df = pd.read_excel(DB_FILE, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        target_order = ['순번', '신품번', '구품번', '품명', '재고수량', '공정 진행중', '입고수량']
        cols = [c for c in target_order if c in df.columns]
        df = df[cols].fillna('')
        return jsonify(df.to_dict(orient='records'))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "파일이 없습니다."}), 400
    
    file = request.files['file']
    filename = file.filename
    
    if not os.path.exists(SAVE_FOLDER):
        os.makedirs(SAVE_FOLDER)
    
    try:
        # CSV 변환 저장 [cite: 751, 756]
        delivery_df = pd.read_excel(file)
        today_str = datetime.now().strftime('%Y%m%d')
        base_name = os.path.splitext(filename)[0]
        csv_filename = f"{base_name}_{today_str}.csv"
        csv_path = os.path.join(SAVE_FOLDER, csv_filename)
        delivery_df.to_csv(csv_path, index=False, encoding='utf-8-sig')

        # 재고 차감 로직 (서식 유지) [cite: 265, 512]
        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i 
                   for i in range(1, ws.max_column + 1)}
        
        delivery_df.columns = [str(c).strip() for c in delivery_df.columns]
        summary = delivery_df.groupby('품번')['납품수량'].sum().reset_index()
        
        chart_data_today = []
        for _, row in summary.iterrows():
            p_no = str(row['품번']).strip()
            qty = row['납품수량']
            chart_data_today.append({"label": p_no, "value": qty})
            
            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                cell_p_no = str(ws.cell(row=r_idx, column=col_map['신품번']).value).strip()
                if cell_p_no == p_no:
                    raw_val = ws.cell(row=r_idx, column=col_map['재고수량']).value
                    current_qty = float(str(raw_val).replace(',', '')) if raw_val else 0
                    ws.cell(row=r_idx, column=col_map['재고수량']).value = current_qty - qty
                    break
        
        wb.save(DB_FILE)
        
        inventory_df = pd.read_excel(DB_FILE, header=2)
        chart_data_total = inventory_df[['신품번', '재고수량']].head(10).to_dict(orient='records')

        return jsonify({
            "status": "success",
            "message": f"{csv_filename} 저장 및 재고 차감 완료",
            "lineData": chart_data_today,
            "barData": chart_data_total
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/update_inventory', methods=['POST'])
def update_inventory():
    """수정 모드 저장 로직 [cite: 167, 766]"""
    try:
        updated_data = request.json
        wb = openpyxl.load_workbook(DB_FILE)
        ws = wb.active
        header_row_num = 3
        
        # ws.max_column 괄호 제거로 에러 방지 [cite: 623, 666]
        col_map = {str(ws.cell(row=header_row_num, column=i).value).strip(): i 
                   for i in range(1, ws.max_column + 1)}

        for item in updated_data:
            p_no = str(item.get('신품번')).strip()
            if not p_no or p_no == 'None': continue

            for r_idx in range(header_row_num + 1, ws.max_row + 1):
                cell_p_no = str(ws.cell(row=r_idx, column=col_map['신품번']).value).strip()
                if cell_p_no == p_no:
                    for col_name, value in item.items():
                        if col_name in col_map and col_name not in ['순번', '신품번']:
                            if col_name in ['재고수량', '공정 진행중', '입고수량']:
                                try:
                                    ws.cell(row=r_idx, column=col_map[col_name]).value = float(str(value).replace(',', ''))
                                except:
                                    ws.cell(row=r_idx, column=col_map[col_name]).value = 0
                            else:
                                ws.cell(row=r_idx, column=col_map[col_name]).value = value
                    break
        
        wb.save(DB_FILE)
        return jsonify({"status": "success", "message": "저장되었습니다."})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def run_flask():
    app.run(host='127.0.0.1', port=5000)

if __name__ == '__main__':
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()
    # pywebview를 이용한 전용 창 실행 [cite: 365, 518]
    webview.create_window('STMS 재고 관리 시스템', 'http://127.0.0.1:5000', width=1200, height=900)
    webview.start()