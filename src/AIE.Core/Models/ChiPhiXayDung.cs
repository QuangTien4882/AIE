namespace AIE.Core.Models;

/// <summary>
/// Chi phí xây dựng (GXD) — kết quả tính toán tổng hợp.
/// Công thức theo Screenshot 2 / TT36/2026:
///   GXD = Gxd + LT
///   Gxd = G + GTGT  (CP XD sau thuế)
///   G   = T + GT + TL  (CP XD trước thuế)
/// </summary>
public class ChiPhiXayDung
{
    // --- Chi phí trực tiếp ---
    /// <summary>Chi phí vật liệu (VL) — thành phần của T</summary>
    public decimal VL { get; set; }

    /// <summary>Chi phí nhân công (NC) — thành phần của T</summary>
    public decimal NC { get; set; }

    /// <summary>Chi phí máy thi công (M) — thành phần của T</summary>
    public decimal M { get; set; }

    /// <summary>T = VL + NC + M — Chi phí trực tiếp</summary>
    public decimal T => VL + NC + M;

    // --- Chi phí gián tiếp ---
    /// <summary>Chi phí chung (CPC) = T × tỉ lệ %</summary>
    public decimal CPC { get; set; }

    /// <summary>CP không XĐ được Khối lượng (TT) = T × tỉ lệ %</summary>
    public decimal TT { get; set; }

    /// <summary>GT = CPC + TT — Chi phí gián tiếp</summary>
    public decimal GT => CPC + TT;

    // --- Thu nhập chịu thuế tính trước ---
    /// <summary>TL = (T + GT) × tỉ lệ % TNCTTT</summary>
    public decimal TL { get; set; }

    // --- Tổng hợp ---
    /// <summary>G = T + GT + TL — CP xây dựng trước thuế</summary>
    public decimal G => T + GT + TL;

    /// <summary>GTGT = G × thuế suất — Thuế giá trị gia tăng</summary>
    public decimal GTGT { get; set; }

    /// <summary>Gxd = G + GTGT — CP xây dựng sau thuế</summary>
    public decimal Gxd => G + GTGT;

    /// <summary>LT = G × tỉ lệ % — Chi phí nhà tạm</summary>
    public decimal LT { get; set; }

    /// <summary>GXD = Gxd + LT — Tổng chi phí xây dựng</summary>
    public decimal GXD => Gxd + LT;

    // --- Tỉ lệ % đã áp dụng (lưu lại để Kiểm tra) ---
    public decimal TiLeCPC { get; set; }
    public decimal TiLeTT { get; set; }
    public decimal TiLeTNCTTT { get; set; }
    public decimal TiLeGTGT { get; set; }
    public decimal TiLeNhaTam { get; set; }
}
