namespace AIE.Core.Models;

/// <summary>
/// Máy thi công — giá Đà Nẵng, cho phép chỉnh sửa.
/// </summary>
public class MayThiCong
{
    public int Id { get; set; }

    /// <summary>Mã máy (VD: M112.1101)</summary>
    public string MaMay { get; set; } = string.Empty;

    /// <summary>Tên máy & công suất (VD: "Máy đầm bê tông, đầm bàn 1.0kW")</summary>
    public string TenMay { get; set; } = string.Empty;

    /// <summary>Đơn vị (ca)</summary>
    public string DonVi { get; set; } = "ca";

    /// <summary>Đơn giá (đồng/ca) — EDITABLE</summary>
    public decimal DonGia { get; set; }

    /// <summary>Ghi chú</summary>
    public string? GhiChu { get; set; }

    public DateTime NgayCapNhat { get; set; } = DateTime.Now;
}
