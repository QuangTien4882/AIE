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
    /// Parse ThanhPhanNhanCong thành danh sách (nhóm, số lượng).
    /// VD: "6:1;3:3;3:1;3:1" -> [(6,1), (3,3), (3,1), (3,1)]
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
            var kv = part.Split(':');
            if (kv.Length == 2 && int.TryParse(kv[0], out int nhom) && decimal.TryParse(kv[1], out decimal sl))
            {
                result.Add((nhom, sl));
            }
        }
        return result;
    }
    
    /// <summary>Tổng số nhân công từ ThanhPhanNhanCong</summary>
    public decimal TongSoNhanCong => GetThanhPhanNhanCong().Sum(x => x.SoLuong);
}

