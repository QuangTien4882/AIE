import re
import json

with open("ocr.txt", "r", encoding="utf-8") as f:
    text = f.read()

# Replace newlines and multiple spaces with a single space
text = re.sub(r'\s+', ' ', text)

# Find all machines
matches = re.finditer(r'(M\d{3}\.\d{4}[a-z]?)\s+(.*?)(?=(?:M\d{3}\.\d{4}[a-z]?)|$)', text)

machines = []
for m in matches:
    ma_may = m.group(1)
    block = m.group(2).strip()
    
    # We want to extract:
    # SoCa (int), KhauHao (float), SuaChua (float), ChiPhiKhac (float)
    # Then optionally Nhiên liệu (float + unit)
    # Then optionally Nhân công (int + text)
    # Then Nguyên giá (float)
    
    # Let's find all numbers and keywords
    # A number can be: 200, 10, 1,80, 5.000, 3, 2,50, 1.100.000
    # Fuel units: lít diezel, lít xăng, kWh
    
    # Using regex to find the sequence of 4 numbers at the start of the data columns
    # Usually it's: (SoCa) (KhauHao) (SuaChua) (ChiPhiKhac)
    # Example: 200 10 1,80 4
    # Wait, there's text before these numbers (the machine name).
    # So we look for: (\d+)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)
    data_match = re.search(r'(\d+)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)\s+(\d+(?:,\d+)?)', block)
    if not data_match:
        continue
        
    so_ca = int(data_match.group(1))
    khau_hao = float(data_match.group(2).replace(',', '.'))
    sua_chua = float(data_match.group(3).replace(',', '.'))
    chi_phi_khac = float(data_match.group(4).replace(',', '.'))
    
    rest = block[data_match.end():].strip()
    
    # Now parse the rest. It could end with NguyenGia
    # Let's extract NguyenGia which is the last number in the string
    # It looks like 5.000 or 1.200.000 or 440
    nguyen_gia_match = re.findall(r'([\d\.]+)', rest)
    nguyen_gia = 0
    if nguyen_gia_match:
        # The last one is nguyen gia
        ng_str = nguyen_gia_match[-1].replace('.', '')
        if ng_str.isdigit():
            nguyen_gia = float(ng_str)
            
    # Now fuel:
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
        
    # Nhan cong
    # Can be: "1 nhân công vận hành", "2 thợ lặn", "1 lái xe"
    # Just look for: (\d+)\s+(nhân công|thợ|lái xe|thuyền)
    nhan_cong = 0
    nc_match = re.search(r'(\d+)\s+(nhân công|thợ|lái|thuyền|nhân)', rest)
    if nc_match:
        nhan_cong = int(nc_match.group(1))
        
    # We also need NhomNhanCong. TT 37 says standard is Nhóm 2, or group mapping. For now default to 2.
    
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
