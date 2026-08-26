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

    /// <summary>Phân loại nhân công (Xây dựng, Vận hành máy, Khác)</summary>
    public AIE.Core.Enums.LoaiNhanCong LoaiNhanCong { get; set; } = AIE.Core.Enums.LoaiNhanCong.XayDung;

    /// <summary>Đơn vị tính (công)</summary>
    public string DonVi { get; set; } = "công";

    /// <summary>Đơn giá Vùng II (đồng/công)</summary>
    public decimal DonGiaVung2 { get; set; }

    /// <summary>Đơn giá Vùng III (đồng/công)</summary>
    public decimal DonGiaVung3 { get; set; }

    /// <summary>Đơn giá Vùng IV (đồng/công)</summary>
    public decimal DonGiaVung4 { get; set; }

    /// <summary>Đơn giá Cù Lao Chàm (đồng/công)</summary>
    public decimal DonGiaCLC { get; set; }

    /// <summary>Ghi chú (VD: "Sở XD ĐN công bố 07/2026")</summary>
    public string? GhiChu { get; set; }

    public DateTime NgayCapNhat { get; set; } = DateTime.Now;

    /// <summary>Backward-compatible property — maps to DonGiaVung2 by default</summary>
    public decimal DonGia
    {
        get => DonGiaVung2;
        set => DonGiaVung2 = value;
    }

    /// <summary>Lấy đơn giá theo vùng</summary>
    public decimal GetDonGia(AIE.Core.Enums.Vung vung)
    {
        return vung switch
        {
            AIE.Core.Enums.Vung.VungII => DonGiaVung2,
            AIE.Core.Enums.Vung.VungIII => DonGiaVung3,
            AIE.Core.Enums.Vung.VungIV => DonGiaVung4,
            AIE.Core.Enums.Vung.CuLaoCham => DonGiaCLC,
            _ => DonGiaVung2
        };
    }

    /// <summary>Cập nhật đơn giá theo vùng</summary>
    public void SetDonGia(AIE.Core.Enums.Vung vung, decimal value)
    {
        switch (vung)
        {
            case AIE.Core.Enums.Vung.VungII: DonGiaVung2 = value; break;
            case AIE.Core.Enums.Vung.VungIII: DonGiaVung3 = value; break;
            case AIE.Core.Enums.Vung.VungIV: DonGiaVung4 = value; break;
            case AIE.Core.Enums.Vung.CuLaoCham: DonGiaCLC = value; break;
            default: DonGiaVung2 = value; break;
        }
    }
}
