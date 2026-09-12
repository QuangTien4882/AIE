namespace AIE.Core.Models;

/// <summary>
/// Chi phí xây dựng (GXD) — kết quả tính toán tổng hợp theo Bảng 3.8 Thông tư 36/2026/TT-BXD:
///   GXDTT = T + GT + TL  (Chi phí xây dựng trước thuế)
///   GTGT  = GXDTT × T_GTGT (Thuế giá trị gia tăng)
///   GXD   = GXDTT + GTGT  (Chi phí xây dựng sau thuế — KHÔNG cộng chi phí nhà tạm LT)
///   LT    = Chi phí nhà tạm để ở và điều hành thi công (mục V Bảng 3.8)
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
    /// <summary>Chi phí chung (C hoặc CPC) = T × tỉ lệ % (hoặc NC × tỉ lệ %)</summary>
    public decimal CPC { get; set; }

    /// <summary>CP không XĐ được Khối lượng (TT) = T × tỉ lệ %</summary>
    public decimal TT { get; set; }

    /// <summary>GT = CPC + TT — Chi phí gián tiếp</summary>
    public decimal GT => CPC + TT;

    // --- Thu nhập chịu thuế tính trước ---
    /// <summary>TL = (T + GT) × tỉ lệ % TNCTTT</summary>
    public decimal TL { get; set; }

    // --- Tổng hợp chi phí xây dựng ---
    /// <summary>GXDTT = T + GT + TL — Chi phí xây dựng trước thuế</summary>
    public decimal GXDTT => T + GT + TL;

    /// <summary>Alias tương thích ngược cho GXDTT</summary>
    public decimal G => GXDTT;

    /// <summary>GTGT = GXDTT × thuế suất — Thuế giá trị gia tăng</summary>
    public decimal GTGT { get; set; }

    /// <summary>GXD = GXDTT + GTGT — Chi phí xây dựng sau thuế (Chuẩn TT 36/2026 Bảng 3.8)</summary>
    public decimal GXD => GXDTT + GTGT;

    /// <summary>Alias tương thích ngược cho GXD</summary>
    public decimal Gxd => GXD;

    /// <summary>LT = Chi phí nhà tạm để ở và điều hành thi công (Mục V Bảng 3.8)</summary>
    public decimal LT { get; set; }

    // --- Tỉ lệ % đã áp dụng (lưu lại để kiểm tra và xuất bảng) ---
    public decimal TiLeCPC { get; set; }
    public decimal TiLeTT { get; set; }
    public decimal TiLeTNCTTT { get; set; }
    public decimal TiLeGTGT { get; set; }
    public decimal TiLeNhaTam { get; set; }

    // --- Ngữ cảnh tính toán ---
    public string GiaiDoan { get; set; } = "Lập dự toán xây dựng";
    public string LoaiCongTrinhNhaTam { get; set; } = "Công trình xây dựng còn lại";
    public bool LaVungSauXa { get; set; } = false;
}

