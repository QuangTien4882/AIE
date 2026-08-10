namespace AIE.Core.Models;

/// <summary>
/// Định mức % quản lý dự án — phụ thuộc loại dự án & quy mô.
/// Nội suy tuyến tính giữa các mốc quy mô.
/// </summary>
public class DinhMucQLDA
{
    public int Id { get; set; }

    /// <summary>Loại dự án (VD: "Dân dụng", "Giao thông",...)</summary>
    public string LoaiDuAn { get; set; } = string.Empty;

    /// <summary>Quy mô tối thiểu (tỷ đồng)</summary>
    public decimal QuyMoMin { get; set; }

    /// <summary>Quy mô tối đa (tỷ đồng, null = không giới hạn)</summary>
    public decimal? QuyMoMax { get; set; }

    /// <summary>Tỉ lệ %</summary>
    public decimal TiLe { get; set; }

    public string? GhiChu { get; set; }
}

/// <summary>
/// Định mức % tư vấn — phụ thuộc loại TV + loại CT + quy mô.
/// </summary>
public class DinhMucTuVan
{
    public int Id { get; set; }

    /// <summary>Loại tư vấn (Thiết kế, Giám sát, Thẩm tra,...)</summary>
    public string LoaiTuVan { get; set; } = string.Empty;

    /// <summary>Loại công trình</summary>
    public string LoaiCongTrinh { get; set; } = string.Empty;

    /// <summary>Quy mô tối thiểu (tỷ đồng)</summary>
    public decimal QuyMoMin { get; set; }

    /// <summary>Quy mô tối đa (tỷ đồng)</summary>
    public decimal? QuyMoMax { get; set; }

    /// <summary>Tỉ lệ %</summary>
    public decimal TiLe { get; set; }

    public string? GhiChu { get; set; }
}
