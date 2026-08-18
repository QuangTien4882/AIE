import fitz
import json
import re

pdf_path = r"C:\Users\quang\.gemini\antigravity-ide\brain\96f85643-e6e2-4978-935e-6c0842c38a46\media__1786438549080.pdf"

doc = fitz.open(pdf_path)
machines = []

for i in range(12, doc.page_count): # 0-indexed, page 13 is index 12
    page = doc[i]
    text = page.get_text("text")
    text = re.sub(r'\s+', ' ', text)
    
    matches = re.finditer(r'(M\d{3}\.\d{4}[a-z]?)\s+(.*?)(?=(?:M\d{3}\.\d{4}[a-z]?)|$)', text)
    for m in matches:
        ma_may = m.group(1)
        block = m.group(2).strip()
        
        # Regex to find numbers representing SoCa, KhauHao, SuaChua, ChiPhiKhac
        # Require a leading space to avoid matching numbers attached to units (like m3)
        data_match = re.search(r'\s+(\d+)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)', block)
        if not data_match:
            continue
            
        so_ca = int(data_match.group(1))
        khau_hao = float(data_match.group(2).replace(',', '.'))
        sua_chua = float(data_match.group(3).replace(',', '.'))
        chi_phi_khac = float(data_match.group(4).replace(',', '.'))
        
        rest = block[data_match.end():].strip()
        
        # Extract NguyenGia
        nguyen_gia_match = re.findall(r'([\d\.]+)', rest)
        nguyen_gia = 0
        if nguyen_gia_match:
            ng_str = nguyen_gia_match[-1].replace('.', '')
            if ng_str.isdigit():
                nguyen_gia = float(ng_str)
                
        # Extract fuel
        xang = 0
        diezel = 0
        dien = 0
        
        if 'lít xăng' in rest:
            fuel_match = re.search(r'([\d,]+)\s+lít xăng', rest)
            if fuel_match: xang = float(fuel_match.group(1).replace(',', '.'))
        if 'lít diezel' in rest:
            fuel_match = re.search(r'([\d,]+)\s+lít diezel', rest)
            if fuel_match: diezel = float(fuel_match.group(1).replace(',', '.'))
        if 'kWh' in rest:
            fuel_match = re.search(r'([\d,]+)\s+kWh', rest)
            if fuel_match: dien = float(fuel_match.group(1).replace(',', '.'))
            
        # Extract NhanCong
        nhan_cong = 0
        nc_match = re.search(r'(\d+)\s+(nhân công|thợ|lái|thuyền|nhân)', rest)
        if nc_match:
            nhan_cong = int(nc_match.group(1))
            
        machines.append({
            "MaMay": ma_may,
            "NguyenGia": nguyen_gia,
            "KhauHao": khau_hao,
            "SuaChua": sua_chua,
            "ChiPhiKhac": chi_phi_khac,
            "DinhMucXang": xang,
            "DinhMucDiezel": diezel,
            "DinhMucDien": dien,
            "SoLuongNhanCong": float(nhan_cong),
            "NhomNhanCong": 2
        })

with open("seed_machines.json", "w", encoding="utf-8") as f:
    json.dump(machines, f, indent=2, ensure_ascii=False)

print(f"Extracted {len(machines)} machines.")
