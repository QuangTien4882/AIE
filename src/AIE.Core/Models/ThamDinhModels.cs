using AIE.Core.Enums;
using System.Collections.Generic;

namespace AIE.Core.Models;

/// <summary>
/// Cấu hình mapping cột cho sheet Đơn Giá Chi Tiết
/// </summary>
public class ThamDinhConfig
{
    public string SheetName { get; set; } = string.Empty;
    public string ColMaHieu { get; set; } = "B";
    public string ColTenCT { get; set; } = "E";
    public string ColDonVi { get; set; } = "F";
    public string ColDinhMuc { get; set; } = "G";
    public string ColDonGia { get; set; } = "H";
    public string ColThanhTien { get; set; } = "I";
    public int DongBatDau { get; set; } = 6;
}

/// <summary>
/// 1 công tác đọc được từ file dự toán cần thẩm định
/// </summary>
public class CongTacThamDinh
{
    public int SoDongExcel { get; set; }
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public List<HaoPhiThamDinh> DanhSachHaoPhi { get; set; } = new();
}

/// <summary>
/// 1 hao phí đọc được từ file dự toán
/// </summary>
public class HaoPhiThamDinh
{
    public int SoDongExcel { get; set; }
    public LoaiHaoPhi Loai { get; set; }
    public string MaHieuHP { get; set; } = string.Empty;
    public string TenHaoPhi { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal DinhMuc { get; set; }
    public decimal DonGia { get; set; }
    public decimal ThanhTien { get; set; } // Nếu cần
}

/// <summary>
/// Kết quả thẩm định 1 công tác
/// </summary>
public class KetQuaCongTacThamDinh
{
    public CongTacThamDinh DuToan { get; set; } = new();
    public CongTacXayDung? DinhMucChuan { get; set; }
    public List<SaiLechDinhMuc> DanhSachSaiLech { get; set; } = new();
    public bool DaDat => DanhSachSaiLech.Count == 0 && DinhMucChuan != null;
}

/// <summary>
/// Chi tiết 1 sai lệch (trên 1 dòng Excel)
/// </summary>
public class SaiLechDinhMuc
{
    public string LoaiLoi { get; set; } = string.Empty;
    public string MoTa { get; set; } = string.Empty;
    public LoaiHaoPhi? LoaiHP { get; set; }
    public int? SoDongExcel { get; set; }
    
    // Dữ liệu dùng để xuất cột
    public HaoPhi? HaoPhiChuan { get; set; }
    public decimal? DonGiaChuan { get; set; } // Lấy từ Repo
    
    // Thuộc tính tiện ích
    public decimal ChenhLechDinhMuc => (HaoPhiChuan != null && HaoPhiDuToan != null) ? HaoPhiDuToan.DinhMuc - HaoPhiChuan.DinhMuc : 0;
    public decimal ChenhLechDonGia => (DonGiaChuan.HasValue && HaoPhiDuToan != null) ? HaoPhiDuToan.DonGia - DonGiaChuan.Value : 0;
    
    // Tính toán Thành tiền thẩm định
    public decimal ThanhTienChuan => (HaoPhiChuan != null && DonGiaChuan.HasValue) ? HaoPhiChuan.DinhMuc * DonGiaChuan.Value : 0;
    public decimal ThanhTienDuToan => (HaoPhiDuToan != null) ? (HaoPhiDuToan.ThanhTien > 0 ? HaoPhiDuToan.ThanhTien : HaoPhiDuToan.DinhMuc * HaoPhiDuToan.DonGia) : 0;
    public decimal ChenhLechThanhTien => ThanhTienDuToan - ThanhTienChuan;
    
    public HaoPhiThamDinh? HaoPhiDuToan { get; set; }
}

/// <summary>
/// Tổng hợp vật tư từ dự toán để so sánh đơn giá
/// </summary>
public class VatTuGiaModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenVatTu { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public LoaiHaoPhi LoaiHP { get; set; }
    public decimal GiaDuToan { get; set; }
    public decimal? GiaChuan { get; set; }
    
    public decimal ChenhLechGia => (GiaChuan.HasValue) ? GiaDuToan - GiaChuan.Value : 0;
}
