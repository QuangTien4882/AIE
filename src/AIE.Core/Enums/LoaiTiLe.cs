namespace AIE.Core.Enums;

/// <summary>
/// Loại tỉ lệ phần trăm trong cấu trúc chi phí xây dựng.
/// Theo Thông tư 36/2026/TT-BXD.
/// </summary>
public enum LoaiTiLe
{
    /// <summary>Chi phí chung (CPC) — tính trên T hoặc NC</summary>
    CPC,
    /// <summary>Chi phí một số công việc không xác định được Khối lượng từ thiết kế (TT) — tính trên T</summary>
    TT,
    /// <summary>Thu nhập chịu thuế tính trước (TNCTTT) — tính trên (T + GT)</summary>
    TNCTTT,
    /// <summary>Thuế giá trị gia tăng (GTGT) — tính trên G</summary>
    GTGT,
    /// <summary>Chi phí nhà tạm để ở và điều hành thi công — tính trên G</summary>
    NhaTam
}
