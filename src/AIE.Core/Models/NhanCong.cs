namespace AIE.Core.Models;

/// <summary>
/// Nhân công xây dựng — giá theo nhóm (1-6), Đà Nẵng, cho phép chỉnh sửa.
/// Theo TT 37/2026/TT-BXD.
/// </summary>
public class NhanCong
{
    public int Id { get; set; }

    /// <summary>Mã nhân công (VD: N97790)</summary>
    public string MaNC { get; set; } = string.Empty;

    /// <summary>Tên (VD: "Nhân công nhóm 2")</summary>
    public string TenNC { get; set; } = string.Empty;

    /// <summary>Nhóm nhân công (1-6 theo TT37/2026)</summary>
    public int Nhom { get; set; }

    /// <summary>Đơn vị tính (công)</summary>
    public string DonVi { get; set; } = "công";

    /// <summary>Đơn giá (đồng/công) — EDITABLE</summary>
    public decimal DonGia { get; set; }

    /// <summary>Ghi chú (VD: "Sở XD ĐN công bố 07/2026")</summary>
    public string? GhiChu { get; set; }

    public DateTime NgayCapNhat { get; set; } = DateTime.Now;
}
