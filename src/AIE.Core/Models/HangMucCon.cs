namespace AIE.Core.Models;

/// <summary>
/// Hạng mục con (Level 2) trong dự toán — cấu phần kỹ thuật / phân đoạn thi công
/// (Ví dụ: Trong Hạng mục Giao thông có: Nền đường, Mặt đường, Hệ thống biển báo...)
/// </summary>
public class HangMucCon
{
    public int STT { get; set; }
    public int RowIndex { get; set; }
    public string MaHangMucCon { get; set; } = string.Empty;
    public string TenHangMucCon { get; set; } = string.Empty;
    
    /// <summary>Danh sách công tác thuộc hạng mục con này</summary>
    public List<DongDuToan> DanhSachCongTac { get; set; } = [];

    /// <summary>Tổng thành tiền Vật liệu của hạng mục con</summary>
    public decimal TongVL => DanhSachCongTac.Sum(x => x.ThanhTienVL);

    /// <summary>Tổng thành tiền Nhân công của hạng mục con</summary>
    public decimal TongNC => DanhSachCongTac.Sum(x => x.ThanhTienNC);

    /// <summary>Tổng thành tiền Máy thi công của hạng mục con</summary>
    public decimal TongMay => DanhSachCongTac.Sum(x => x.ThanhTienMay);

    /// <summary>Tổng chi phí trực tiếp của hạng mục con T = VL + NC + M</summary>
    public decimal TongChiPhi => TongVL + TongNC + TongMay;
}
