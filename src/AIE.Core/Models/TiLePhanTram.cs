using AIE.Core.Enums;

namespace AIE.Core.Models;

/// <summary>
/// Tỉ lệ phần trăm — phụ thuộc loại công trình.
/// VD: CPC = 7.3% cho Dân dụng, TNCTTT = 5.5%,...
/// Theo Phụ lục TT 36/2026/TT-BXD.
/// </summary>
public class TiLePhanTram
{
    public int Id { get; set; }

    /// <summary>Loại công trình (Dân dụng, Giao thông,...)</summary>
    public string LoaiCongTrinh { get; set; } = string.Empty;

    /// <summary>Cấp công trình (null nếu không phân cấp)</summary>
    public string? CapCongTrinh { get; set; }

    /// <summary>Loại tỉ lệ (CPC, TT, TNCTTT, GTGT, NhaTam)</summary>
    public string LoaiTiLe { get; set; } = string.Empty;

    /// <summary>Giá trị tỉ lệ % (VD: 7.3, 2.5, 5.5, 8, 1.1)</summary>
    public decimal GiaTri { get; set; }

    /// <summary>Cơ sở tính (T, T+GT, G)</summary>
    public string CoSoTinh { get; set; } = string.Empty;
}
