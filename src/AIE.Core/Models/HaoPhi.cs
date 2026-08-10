using AIE.Core.Enums;

namespace AIE.Core.Models;

/// <summary>
/// Hao phí (VL/NC/Máy) của 1 công tác xây dựng.
/// VD: Xi măng PCB40 - V08770 - 234.725 kg cho mã hiệu AF.11112.
/// </summary>
public class HaoPhi
{
    public int Id { get; set; }
    public int CongTacId { get; set; }

    /// <summary>VL / NC / MAY</summary>
    public LoaiHaoPhi LoaiHaoPhi { get; set; }

    /// <summary>Mã hiệu hao phí (VD: V08770, N97790, M112.1101)</summary>
    public string MaHieuHP { get; set; } = string.Empty;

    /// <summary>Tên hao phí (VD: "Xi măng PCB40", "Nhân công nhóm 2")</summary>
    public string TenHaoPhi { get; set; } = string.Empty;

    /// <summary>Đơn vị hao phí (kg, m3, công, ca,...)</summary>
    public string DonVi { get; set; } = string.Empty;

    /// <summary>Mức hao phí định mức (VD: 234.725)</summary>
    public decimal DinhMuc { get; set; }

    /// <summary>Hệ số (mặc định = 1.0)</summary>
    public decimal HeSo { get; set; } = 1.0m;
}
