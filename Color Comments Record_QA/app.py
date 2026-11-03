# -*- coding: utf-8 -*-
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import os
import shutil

APP_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
CORS(app)

HTML_FILE = os.path.join(APP_DIR, 'Color Comments Record_QA.html')

@app.route('/')
def index():
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        return f.read(), 200, {'Content-Type': 'text/html; charset=utf-8'}

@app.route('/api/save', methods=['POST'])
def save_excel():
    data = request.get_json()
    action = data.get('action', 'add')
    record = data.get('record', {})
    # 如果沒有提供 filePath，使用硬編碼的預設檔案（用於測試）
    file_path = data.get('filePath', 'Color Comments Record_QA.xlsx')
    print(f"[DEBUG] 收到的 filePath: {file_path}")

    if not os.path.isabs(file_path):
        file_path = os.path.join(APP_DIR, file_path)

    if not os.path.exists(file_path):
        return jsonify({'success': False, 'message': 'File not found'}), 400

    shutil.copy2(file_path, file_path + '.bak')

    wb = load_workbook(file_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]

    qa_name_code = {'Amber': 'A', 'Asther': 'B', 'Rick': 'C', 'Vicky': 'D'}
    sample_type_code = {'L/S/H': 'A', '1ST LOT': 'B'}

    qa_code = qa_name_code.get(record.get('核色人員', ''), '?')
    year = '26'
    # 編號序號：掃描現有編號，排序，檢查起始點，確保從 00001 開始
    編號_col = headers.index('編號') + 1 if '編號' in headers else None
    existing_seqs = []
    if 編號_col:
        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=編號_col).value
            if cell_val and isinstance(cell_val, str) and f"{year}-" in cell_val:
                try:
                    seq_str = cell_val.split('-')[1]  # 提取中間的序號
                    seq_num = int(seq_str)
                    existing_seqs.append(seq_num)
                except:
                    pass

    # 排序並找出下一個可用序號
    if existing_seqs:
        existing_seqs.sort()
        global_seq = str(existing_seqs[-1] + 1).zfill(5)  # 最高序號 + 1
    else:
        global_seq = '00001'  # 無現存編號時從 00001 開始
    sample_type = record.get('樣品類別', '')
    sample_code = sample_type_code.get(sample_type, '?')
    # 樣品序號：同樣排序後檢查，確保連續
    樣品類別_col = headers.index('樣品類別') + 1 if '樣品類別' in headers else None
    sample_seqs = []
    if 樣品類別_col:
        for row in range(2, ws.max_row + 1):
            cell_type = ws.cell(row=row, column=樣品類別_col).value
            if cell_type == sample_type:
                # 從同行編號提取樣品序號（編號格式：QA-YY-XXXXX-ST）
                編號_cell = ws.cell(row=row, column=編號_col).value if 編號_col else None
                if 編號_cell and isinstance(編號_cell, str):
                    try:
                        sample_part = 編號_cell.split('-')[-1]  # 最後一部分
                        seq_num = int(sample_part[1:])  # 去掉樣品代碼，只取數字
                        sample_seqs.append(seq_num)
                    except:
                        pass

    if sample_seqs:
        sample_seqs.sort()
        sample_seq = str(sample_seqs[-1] + 1).zfill(2)
    else:
        sample_seq = '01'

    record_no = f"{qa_code}{year}-{global_seq}-{sample_code}{sample_seq}"
    record['編號'] = record_no

    # 自動生成收件序號：年份 + 遞增序號
    if '收件序號' in headers:
        col_idx = headers.index('收件序號') + 1
        max_seq = 0
        for row in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=col_idx).value
            if cell_val and isinstance(cell_val, str) and cell_val.startswith(year):
                try:
                    seq_num = int(cell_val[2:])  # 提取 year 後的數字
                    max_seq = max(max_seq, seq_num)
                except:
                    pass
        next_seq = max_seq + 1
        record['收件序號'] = f"{year}{str(next_seq).zfill(4)}"


    # 自動填入預定天數（如果未提供）
    if not record.get('預定天數'):
        record['預定天數'] = 1.5

    # 自動計算實際天數（工作天數，以半天為單位）
    receive_date_str = record.get('收件日', '')
    ship_date_str = record.get('出件日', '')
    receive_period = record.get('時段', '').lower() if record.get('時段') else 'am'
    ship_period = record.get('時段', '').lower() if record.get('時段') else 'pm'

    if receive_date_str and ship_date_str:
        try:
            from datetime import datetime, timedelta
            recv_date_only = receive_date_str.split('T')[0] if 'T' in receive_date_str else receive_date_str
            ship_date_only = ship_date_str.split('T')[0] if 'T' in ship_date_str else ship_date_str

            recv_dt = datetime.strptime(recv_date_only, '%Y-%m-%d')
            ship_dt = datetime.strptime(ship_date_only, '%Y-%m-%d')

            # 計算工作天數（不含週末）
            work_days = 0
            days_diff = (ship_dt - recv_dt).days
            for i in range(days_diff + 1):
                current_day = recv_dt + timedelta(days=i)
                if current_day.weekday() < 5:  # 0-4 是平日，5-6 是週末
                    work_days += 1

            actual_days = float(work_days)
            if receive_period == 'am':
                actual_days -= 0.5
            elif receive_period == 'pm':
                actual_days -= 1

            if ship_period == 'am':
                actual_days -= 0.5

            record['實際天數'] = max(0.5, actual_days)
        except:
            pass

    # 根據日期自動填時段（時段欄位在日期欄位之後）
    receive_idx = headers.index('收件日') if '收件日' in headers else -1
    ship_idx = headers.index('出件日') if '出件日' in headers else -1

    if receive_idx >= 0 and receive_idx + 1 < len(headers):
        if headers[receive_idx + 1] == '時段':
            receive_date = record.get('收件日', '')
            if receive_date and 'T' in receive_date:
                hour = int(receive_date.split('T')[1].split(':')[0])
                record['收件日_時段'] = 'am' if hour < 12 else 'pm'

    if ship_idx >= 0 and ship_idx + 1 < len(headers):
        if headers[ship_idx + 1] == '時段':
            ship_date = record.get('出件日', '')
            if ship_date and 'T' in ship_date:
                hour = int(ship_date.split('T')[1].split(':')[0])
                record['出件日_時段'] = 'am' if hour < 12 else 'pm'

    processed_record = {}
    for i, header in enumerate(headers):
        value = record.get(header, '')
        # 只對日期欄位轉換格式（移除時間部分）
        if header in ['收件日', '出件日', 'CQ上傳日期']:
            if isinstance(value, str) and 'T' in value:
                date_part = value.split('T')[0].replace('-', '/')
                value = date_part
            elif isinstance(value, str) and '-' in value and len(value) == 10:
                value = value.replace('-', '/')
        # 時段欄位：根據前一個日期欄位判定（區分收件/出件時段）
        if header == '時段' and i > 0:
            prev_header = headers[i - 1]
            prev_value = record.get(prev_header, '')

            # 如果用戶已經選擇了時段（am/pm），就保持
            if value and value.lower() in ['am', 'pm']:
                pass
            else:
                # 自動判定：根據前一個日期欄位的時間
                if (prev_header in ['收件日', '出件日'] and
                    isinstance(prev_value, str) and 'T' in prev_value):
                    try:
                        hour = int(prev_value.split('T')[1].split(':')[0])
                        value = 'am' if hour < 12 else 'pm'
                    except Exception as e:
                        print(f"[ERROR] 時段判定失敗: {e}")

            # 根據前一欄位區分是收件時段還是出件時段
            if prev_header == '收件日':
                processed_record['收件時段'] = value
            elif prev_header == '出件日':
                processed_record['出件時段'] = value
            else:
                processed_record[header] = value
        else:
            processed_record[header] = value

    if action == 'add':
        # 找最後一筆有數據的行，新記錄寫在下一行
        # 從後往前掃描，找到第一個有數據的行
        last_data_row = 1  # 至少有標題行
        編號_col = headers.index('編號') + 1 if '編號' in headers else None

        for row in range(ws.max_row, 1, -1):
            if 編號_col and ws.cell(row=row, column=編號_col).value:
                last_data_row = row
                break

        new_row = last_data_row + 1
        print(f"[DEBUG] 保存前 - 最後數據行: {last_data_row}, 新記錄寫入第 {new_row} 行")

        for col_idx, header in enumerate(headers, start=1):
            val = processed_record.get(header, '')
            # 特殊處理時段欄位：根據前一欄位決定用哪個時段
            if header == '時段' and col_idx > 1:
                prev_header = headers[col_idx - 2] if col_idx > 1 else ''
                if prev_header == '收件日':
                    val = processed_record.get('收件時段', '')
                elif prev_header == '出件日':
                    val = processed_record.get('出件時段', '')
            cell = ws.cell(row=new_row, column=col_idx, value=val)
            cell.number_format = '@'

        print(f"[DEBUG] 保存後 - 編號: {processed_record.get('編號', '')}, 行號: {new_row}")

    elif action == 'update':
        match_fields = ['客戶', '申請人', 'Style No', '供應商', '品項', '品名', '染法', '顏色/花版']
        found = False

        for row_idx in range(2, ws.max_row + 1):
            match = True
            for field in match_fields:
                col_idx = headers.index(field) + 1 if field in headers else None
                if col_idx:
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    record_value = processed_record.get(field, '')
                    if str(cell_value or '') != str(record_value or ''):
                        match = False
                        break

            if match:
                for col_idx, header in enumerate(headers, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=processed_record.get(header, ''))
                    cell.number_format = '@'
                found = True
                break

        if not found:
            new_row = ws.max_row + 1
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=new_row, column=col_idx, value=processed_record.get(header, ''))
                cell.number_format = '@'

    wb.save(file_path)
    wb.close()

    return jsonify({'success': True, 'message': 'OK'}), 200

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
