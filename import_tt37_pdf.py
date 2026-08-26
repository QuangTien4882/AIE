import re
import os
import sqlite3

def clean_float(s):
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def process_block(lines):
    if len(lines) < 8: return None # A real machine has at least 8 lines
    
    ma_may = lines[0].strip()
    if ma_may.endswith('00') and len(lines) < 5: return None # Likely a group heading
    
    if len(lines) < 8: return None
    
    # Find the 4 numeric values: SoCaNam, KhauHao, SuaChua, ChiPhiKhac
    # SoCaNam is usually >= 100.
    start_idx = -1
    for i in range(1, len(lines) - 4):
        # Check if lines i, i+1, i+2, i+3 are numeric
        def is_num(s):
            return re.match(r'^[\d,\.]+$', s.strip()) is not None
        
        if is_num(lines[i]) and is_num(lines[i+1]) and is_num(lines[i+2]) and is_num(lines[i+3]):
            so_ca = clean_float(lines[i])
            if so_ca >= 50: # Some machines have SoCaNam = 90
                start_idx = i
                break
                
    if start_idx == -1: return None
    
    so_ca = int(clean_float(lines[start_idx]))
    khau_hao = clean_float(lines[start_idx + 1])
    sua_chua = clean_float(lines[start_idx + 2])
    chi_phi_khac = clean_float(lines[start_idx + 3])
    
    # Find NguyenGia from the end
    nguyen_gia = 0
    end_idx = len(lines) - 1
    for i in range(len(lines) - 1, start_idx + 3, -1):
        try:
            val = float(lines[i].replace('.', '').replace(',', '.'))
            if val >= 1000 or ('.' in lines[i] and val >= 1): 
                # NguyenGia is always large, or has a dot like "9.000"
                nguyen_gia = val * 1000
                end_idx = i
                break
        except:
            pass
            
    middle_lines = lines[start_idx + 4 : end_idx]
    middle_text = " ".join(middle_lines)
    
    # Remove known garbage headers from middle_text
    garbage = [
        "Mã hiệu Loại máy và thiết bị Số ca năm Định mức (%)",
        "Mã hiệu Loại máy và thiết bị",
        "Số ca năm",
        "Định mức (%)",
        "Định mức tiêu hao nhiên liệu, năng lượng (1ca)",
        "Định mức tiêu", "hao nhiên liệu,", "năng lượng", "(1ca)",
        "Nhân công vận hành, điều khiển máy",
        "Nguyên giá tham khảo (1.000 VND)", "Nguyên giá", "tham khảo", "(1.000 VND)",
        "Khấu hao Sửa chữa Chi phí khác",
        "Khấu hao", "Sửa chữa", "Chi phí khác",
        "1 2 3 4 5 6 7 8 9",
        "1 2 3 4 5 6 7 8",
        "BẢNG ĐỊNH MỨC CÁC HAO PHÍ XÁC ĐỊNH GIÁ CA MÁY VÀ THIẾT BỊ THI CÔNG XÂY DỰNG"
    ]
    for g in garbage:
        middle_text = middle_text.replace(g, "")
    
    # Clean up single digits that are orphaned headers
    # The header digits are usually separated by spaces like "1 2 3 4 5 6 7 8"
    middle_text = re.sub(r'\b[1-9]\b(?=\s+[1-9]\b)', '', middle_text)
    
    # Extract fuels
    xang = 0
    diezel = 0
    dien = 0
    
    m_xang = re.search(r'([\d,\.]+)\s*lít xăng', middle_text)
    if m_xang: xang = clean_float(m_xang.group(1))
    
    m_diezel = re.search(r'([\d,\.]+)\s*lít diezel', middle_text)
    if m_diezel: diezel = clean_float(m_diezel.group(1))
    
    m_dien = re.search(r'([\d,\.]+)\s*kWh', middle_text)
    if m_dien: dien = clean_float(m_dien.group(1))
    
    # Remove fuels to get nhan cong
    nc_str = middle_text
    nc_str = re.sub(r'[\d,\.]+\s*lít xăng', '', nc_str)
    nc_str = re.sub(r'[\d,\.]+\s*lít diezel', '', nc_str)
    nc_str = re.sub(r'[\d,\.]+\s*kWh', '', nc_str)
    nc_str = nc_str.replace('+', ' + ').strip()
    nc_str = re.sub(r'\s+', ' ', nc_str)
    
    # Remove header column digits
    nc_str = nc_str.replace("1 2 3 4 5 6 7 8 9", "")
    nc_str = nc_str.replace("1 2 3 4 5 6 7 8", "")
    
    if nc_str.startswith('+'): nc_str = nc_str[1:].strip()
    if nc_str.endswith('+'): nc_str = nc_str[:-1].strip()
    
    # Build ThanhPhanNhanCong
    tp_parts = []
    for role_part in nc_str.split('+'):
        role_part = role_part.strip()
        if not role_part: continue
        
        qty_match = re.match(r'^([\d,\.]+)\s+(.*)', role_part)
        if qty_match:
            qty = clean_float(qty_match.group(1))
            role_name = qty_match.group(2).lower()
            
            nhom = 1
            if 'lái xe' in role_name: nhom = 2
            elif 'thuyền trưởng' in role_name: nhom = 4
            elif 'thợ lặn' in role_name: nhom = 5
            elif 'kỹ sư' in role_name or 'kỹ thuật' in role_name: nhom = 6
            elif 'thợ máy' in role_name or 'thợ điện' in role_name or 'thuỷ thủ' in role_name or 'thuyền phó' in role_name: nhom = 3
            
            tp_parts.append(f"{nhom}:{qty:g}")
            
    thanh_phan = ";".join(tp_parts)
    
    return {
        'MaMay': ma_may,
        'NguyenGia': nguyen_gia,
        'KhauHao': khau_hao,
        'SuaChua': sua_chua,
        'ChiPhiKhac': chi_phi_khac,
        'Xang': xang,
        'Diezel': diezel,
        'Dien': dien,
        'SoCaNam': so_ca,
        'NhanCongString': nc_str,
        'ThanhPhanNhanCong': thanh_phan
    }

def main():
    text_path = r"C:\Users\quang\.gemini\antigravity-ide\brain\96f85643-e6e2-4978-935e-6c0842c38a46\scratch\pdf_text.txt"
    with open(text_path, 'r', encoding='utf-8') as f:
        lines = [l.strip() for l in f]
        
    blocks = []
    current_block = []
    
    for line in lines:
        if not line: continue
        
        if re.match(r'^M\d{3}\.\d{4}[a-zA-Z]?$', line):
            if current_block:
                blocks.append(current_block)
            current_block = [line]
        elif current_block:
            current_block.append(line)
            
    if current_block:
        blocks.append(current_block)
        
    machines = []
    for b in blocks:
        m = process_block(b)
        if m: machines.append(m)
        
    print(f"Parsed {len(machines)} machines.")
    
    # Update Database
    db_path = os.path.join(os.environ['LOCALAPPDATA'], 'AIE_DuToan', 'aie_database.sqlite')
    if not os.path.exists(db_path):
        print("DB not found at", db_path)
        return
        
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    updated_count = 0
    for m in machines:
        # We only UPDATE existing machines, or we INSERT OR REPLACE?
        # The user wants "cập nhật dữ liệu từ PDF", so we can UPDATE.
        # But maybe they are not in the DB yet? We should use INSERT OR REPLACE.
        # Wait, if we use INSERT OR REPLACE, we might overwrite TenMay if there was a join table, but DinhMucCaMay_TT37 doesn't have TenMay.
        # So INSERT OR REPLACE is safe.
        
        # Check if machine exists
        c.execute("SELECT MaMay FROM DinhMucCaMay_TT37 WHERE MaMay=?", (m['MaMay'],))
        exists = c.fetchone()
        
        sql = """
            INSERT OR REPLACE INTO DinhMucCaMay_TT37 
            (MaMay, NguyenGia, KhauHao, SuaChua, ChiPhiKhac, DinhMucXang, DinhMucDiezel, DinhMucDien, SoCaNam, NhanCongString, ThanhPhanNhanCong, SoLuongNhanCong, NhomNhanCong)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
        """
        c.execute(sql, (
            m['MaMay'], m['NguyenGia'], m['KhauHao'], m['SuaChua'], m['ChiPhiKhac'], 
            m['Xang'], m['Diezel'], m['Dien'], m['SoCaNam'], m['NhanCongString'], m['ThanhPhanNhanCong']
        ))
        updated_count += 1
        
    conn.commit()
    conn.close()
    
    print(f"Updated {updated_count} machines in database.")

if __name__ == "__main__":
    main()
