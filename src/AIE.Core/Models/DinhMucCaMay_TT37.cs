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
    
    /// <summary>Số lượng nhân công điều khiển máy (người/ca)</summary>
    public decimal SoLuongNhanCong { get; set; }
    
    /// <summary>Nhóm nhân công (1, 2, 3, 4) để ánh xạ lấy đơn giá nhóm tương ứng</summary>
    public int NhomNhanCong { get; set; }
}
