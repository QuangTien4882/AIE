namespace AIE.Core.Models;

/// <summary>
/// Công tác xây dựng — đại diện cho 1 mã hiệu định mức (VD: AF.11112).
/// Universal — không phân biệt lĩnh vực/loại công trình.
/// </summary>
public class CongTacXayDung
{
    public int Id { get; set; }

    /// <summary>Mã hiệu định mức (VD: AF.11112, AB.11563)</summary>
    public string MaHieu { get; set; } = string.Empty;

    /// <summary>Tên công tác (VD: "Bê tông lót móng SX bằng máy trộn...")</summary>
    public string TenCongTac { get; set; } = string.Empty;

    /// <summary>Đơn vị tính (VD: m3, 100m3, tấn, m2,...)</summary>
    public string DonVi { get; set; } = string.Empty;

    /// <summary>Danh sách hao phí VL/NC/Máy</summary>
    public List<HaoPhi> DanhSachHaoPhi { get; set; } = [];
}
