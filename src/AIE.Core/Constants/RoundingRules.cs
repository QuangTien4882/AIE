namespace AIE.Core.Constants;

/// <summary>
/// Quy tắc làm tròn số trong dự toán xây dựng (Thông tư 36/37/38).
/// </summary>
public static class RoundingRules
{
    /// <summary>Làm tròn thành tiền/chi phí (đồng) — lấy phần nguyên</summary>
    public const int ThanhTien = 0;

    /// <summary>Làm tròn đơn giá vật liệu/nhân công/máy (đồng) — lấy phần nguyên</summary>
    public const int DonGia = 0;

    /// <summary>Làm tròn khối lượng công tác — 3 chữ số thập phân</summary>
    public const int KhoiLuong = 3;

    /// <summary>Làm tròn định mức hao phí — 3 đến 5 chữ số thập phân</summary>
    public const int HaoPhi = 5;

    /// <summary>Làm tròn định mức tỉ lệ phần trăm (%) — 3 chữ số thập phân</summary>
    public const int TiLePhanTram = 3;

    /// <summary>Làm tròn chỉ số giá — 2 chữ số thập phân</summary>
    public const int ChiSoGia = 2;
}
