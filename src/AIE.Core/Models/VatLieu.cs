namespace AIE.Core.Models;

/// <summary>
/// Vật liệu xây dựng — giá Đà Nẵng, cho phép chỉnh sửa.
/// </summary>
public class VatLieu
{
    public int Id { get; set; }

    /// <summary>Mã vật liệu (VD: V08770) — khớp với HaoPhi.MaHieuHP</summary>
    public string MaVL { get; set; } = string.Empty;

    /// <summary>Tên vật liệu (VD: "Xi măng PCB40")</summary>
    public string TenVL { get; set; } = string.Empty;

    /// <summary>Đơn vị tính (kg, m3, lít,...)</summary>
    public string DonVi { get; set; } = string.Empty;

    /// <summary>Đơn giá (đồng) — EDITABLE</summary>
    public decimal DonGia { get; set; }

    /// <summary>Cước vận chuyển (đồng) — EDITABLE</summary>
    public decimal CuocVanChuyen { get; set; }

    /// <summary>Nhà sản xuất / Nguồn gốc (optional)</summary>
    public string? NhaSanXuat { get; set; }

    /// <summary>Ghi chú (VD: "Giá Liên Sở Q3/2026")</summary>
    public string? GhiChu { get; set; }

    public DateTime NgayCapNhat { get; set; } = DateTime.Now;
}
