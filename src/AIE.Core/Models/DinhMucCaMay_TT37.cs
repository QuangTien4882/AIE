using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models;

public class DinhMucCaMay_TT37
{
    public string MaMay { get; set; } = string.Empty;
    public decimal NguyenGia { get; set; }
    
    /// <summary>Tỷ lệ % khấu hao năm</summary>
    public decimal KhauHao { get; set; }
    
    /// <summary>Tỷ lệ % sửa chữa năm</summary>
    public decimal SuaChua { get; set; }
    
    /// <summary>Tỷ lệ % chi phí khác năm</summary>
    public decimal ChiPhiKhac { get; set; }
    
    /// <summary>Định mức tiêu hao xăng (lít/ca)</summary>
    public decimal DinhMucXang { get; set; }
    
    /// <summary>Định mức tiêu hao diezel (lít/ca)</summary>
    public decimal DinhMucDiezel { get; set; }
    
    /// <summary>Định mức tiêu hao điện (kWh/ca)</summary>
    public decimal DinhMucDien { get; set; }
    
    /// <summary>Số lượng nhân công điều khiển máy (người/ca) - legacy single group</summary>
    public decimal SoLuongNhanCong { get; set; }
    
    /// <summary>Nhóm nhân công (1, 2, 3, 4) - legacy single group</summary>
    public int NhomNhanCong { get; set; }
    
    /// <summary>Số ca làm việc của máy trong năm (ca/năm). Mặc định 250 nếu chưa có dữ liệu</summary>
    public int SoCaNam { get; set; } = 250;
    
    /// <summary>Chuỗi mô tả nhân công gốc từ TT37 (VD: "1 thuyền phó + 3 thợ máy + 1 thợ điện + 1 thủy thủ")</summary>
    public string NhanCongString { get; set; }
    
    /// <summary>Thành phần nhân công cấu trúc (VD: "6:1;3:3;3:1;3:1" = nhóm:số_lượng)</summary>
    public string ThanhPhanNhanCong { get; set; }
    
    // === Hệ số nhiên liệu phụ theo TT37/2026 ===
    public const decimal HeSoNLPhu_Xang = 1.02m;
    public const decimal HeSoNLPhu_Diezel = 1.03m;
    public const decimal HeSoNLPhu_Dien = 1.05m;
    
    /// <summary>Hệ số nhiên liệu phụ cho máy này (dựa vào loại nhiên liệu chính)</summary>
    public decimal HeSoNhienLieuPhu
    {
        get
        {
            if (DinhMucXang > 0) return HeSoNLPhu_Xang;
            if (DinhMucDiezel > 0) return HeSoNLPhu_Diezel;
            if (DinhMucDien > 0) return HeSoNLPhu_Dien;
            return 1.0m;
        }
    }
    /// <summary>
    /// Mapping mã chức danh sang nhóm nhân công vận hành (II-x).
    /// Dùng khi ThanhPhanNhanCong lưu theo format mới "LX:1;TD:2"
    /// </summary>
    private static readonly Dictionary<string, int> ChucDanhToNhom = new()
    {
        { "VH", 1 },  { "VH1", 1 },  // Nhân công vận hành máy (backward compat)
        { "LX", 2 },                  // Lái xe
        { "TT", 3 },  { "TD", 3 },   // Thủy thủ (backward compat)
        { "TM", 3 },                  // Thợ máy
        { "TDIEN", 3 },               // Thợ điện
        { "MTR", 4 },  { "MT", 4 },   // Máy trưởng (backward compat)
        { "MAY2", 4 },                // Máy II
        { "DTRU", 4 },                // Điện trưởng
        { "TTRU", 4 },                // Thuyền trưởng
        { "TPHO", 4 },                // Thuyền phó
        { "KTV1", 4 },                // Kỹ thuật viên cuốc I
        { "KTV2", 4 },                // Kỹ thuật viên cuốc II
        { "MAY1", 4 },                // Máy I
    };

    /// <summary>
    /// Parse ThanhPhanNhanCong thành danh sách (nhóm, số lượng).
    /// Hỗ trợ cả format cũ "6:1;3:3" và format mới "LX:1;TD:2"
    /// </summary>
    public List<(int Nhom, decimal SoLuong)> GetThanhPhanNhanCong()
    {
        var result = new List<(int, decimal)>();
        if (string.IsNullOrWhiteSpace(ThanhPhanNhanCong))
        {
            // Fallback to legacy single group
            if (NhomNhanCong > 0 && SoLuongNhanCong > 0)
            {
                int multiplier = 1;
                if (NhomNhanCong == 3) multiplier = 3; // "thủy thủ, thợ máy, thợ điện"
                else if (NhomNhanCong == 4 || NhomNhanCong == 6) multiplier = 2; // "máy trưởng, thuyền trưởng" or "thuyền trưởng, thuyền phó"
                
                result.Add((NhomNhanCong, SoLuongNhanCong * multiplier));
            }
            return result;
        }
        
        foreach (var part in ThanhPhanNhanCong.Split(';'))
        {
            var kv = part.Trim().Split(':');
            if (kv.Length != 2) continue;
            
            if (!decimal.TryParse(kv[1], System.Globalization.NumberStyles.Any, 
                System.Globalization.CultureInfo.InvariantCulture, out decimal sl)) continue;
            
            string key = kv[0].Trim();
            
            if (int.TryParse(key, out int nhom))
            {
                // Format cũ: số nhóm trực tiếp
                result.Add((nhom, sl));
            }
            else if (ChucDanhToNhom.TryGetValue(key, out int nhomMapped))
            {
                // Format mới: mã chức danh
                result.Add((nhomMapped, sl));
            }
        }
        return result;
    }
    
    /// <summary>Tổng số nhân công từ ThanhPhanNhanCong</summary>
    public decimal TongSoNhanCong => GetThanhPhanNhanCong().Sum(x => x.SoLuong);
}

