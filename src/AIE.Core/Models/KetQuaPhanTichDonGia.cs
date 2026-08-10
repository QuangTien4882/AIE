namespace AIE.Core.Models;

/// <summary>
/// Kết quả phân tích đơn giá chi tiết cho 1 công tác (giống Screenshot 2).
/// Tính đơn giá VL + NC + M dựa trên hao phí × đơn giá địa phương.
/// </summary>
public class KetQuaPhanTichDonGia
{
    /// <summary>Mã hiệu định mức</summary>
    public string MaHieu { get; set; } = string.Empty;

    /// <summary>Tên công tác</summary>
    public string TenCongTac { get; set; } = string.Empty;

    /// <summary>Đơn vị</summary>
    public string DonVi { get; set; } = string.Empty;

    /// <summary>Chi tiết hao phí VL với đơn giá & thành tiền</summary>
    public List<ChiTietHaoPhi> ChiTietVatLieu { get; set; } = [];

    /// <summary>Chi tiết hao phí NC với đơn giá & thành tiền</summary>
    public List<ChiTietHaoPhi> ChiTietNhanCong { get; set; } = [];

    /// <summary>Chi tiết hao phí Máy với đơn giá & thành tiền</summary>
    public List<ChiTietHaoPhi> ChiTietMay { get; set; } = [];

    /// <summary>Tổng đơn giá vật liệu (đồng)</summary>
    public decimal DonGiaVL => ChiTietVatLieu.Sum(x => x.ThanhTien);

    /// <summary>Tổng đơn giá nhân công (đồng)</summary>
    public decimal DonGiaNC => ChiTietNhanCong.Sum(x => x.ThanhTien);

    /// <summary>Tổng đơn giá máy (đồng)</summary>
    public decimal DonGiaMay => ChiTietMay.Sum(x => x.ThanhTien);

    /// <summary>Đơn giá tổng hợp = VL + NC + M</summary>
    public decimal DonGiaTongHop => DonGiaVL + DonGiaNC + DonGiaMay;
}

/// <summary>
/// 1 dòng chi tiết hao phí trong bảng PTĐG.
/// VD: V08770 | Xi măng PCB40 | kg | 234.725 | 1,741 | 408,656
/// </summary>
public class ChiTietHaoPhi
{
    /// <summary>Mã hiệu hao phí (V08770, N97790, M112.1101)</summary>
    public string MaHieuHP { get; set; } = string.Empty;

    /// <summary>Tên hao phí</summary>
    public string TenHaoPhi { get; set; } = string.Empty;

    /// <summary>Đơn vị</summary>
    public string DonVi { get; set; } = string.Empty;

    /// <summary>Định mức hao phí</summary>
    public decimal DinhMuc { get; set; }

    /// <summary>Đơn giá (đồng) — từ bảng giá Đà Nẵng</summary>
    public decimal DonGia { get; set; }

    /// <summary>Hệ số</summary>
    public decimal HeSo { get; set; } = 1.0m;

    /// <summary>Thành tiền = Định mức × Đơn giá × Hệ số</summary>
    public decimal ThanhTien => DinhMuc * DonGia * HeSo;
}
