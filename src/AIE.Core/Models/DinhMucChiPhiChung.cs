namespace AIE.Core.Models;

/// <summary>
/// Định mức Chi phí chung (Bảng 3.3 Thông tư 36/2026/TT-BXD)
/// </summary>
public class DinhMucCPC
{
    public int Id { get; set; }
    
    /// <summary>Loại công trình (Dân dụng, Công nghiệp,...)</summary>
    public string LoaiCongTrinh { get; set; } = string.Empty;
    
    /// <summary>Phân loại phụ (Tu bổ di tích, Hầm lò,...)</summary>
    public string? PhanLoaiPhu { get; set; }
    
    /// <summary>Quy mô chi phí XD tối thiểu (tỷ đồng)</summary>
    public decimal QuyMoMin { get; set; }
    
    /// <summary>Quy mô chi phí XD tối đa (tỷ đồng, null = > 1000)</summary>
    public decimal? QuyMoMax { get; set; }
    
    /// <summary>Tỷ lệ % chi phí chung</summary>
    public decimal TiLe { get; set; }
}

/// <summary>
/// Định mức Chi phí một số công việc không xác định được KL từ TK (Bảng 3.5 Thông tư 36/2026/TT-BXD)
/// </summary>
public class DinhMucTT
{
    public int Id { get; set; }
    
    /// <summary>Loại công trình (Dân dụng, Công nghiệp,...)</summary>
    public string LoaiCongTrinh { get; set; } = string.Empty;
    
    /// <summary>Phân loại phụ (Tu bổ di tích, Hầm lò,...)</summary>
    public string? PhanLoaiPhu { get; set; }
    
    /// <summary>Tỷ lệ % chi phí không xác định được KL</summary>
    public decimal TiLe { get; set; }
}
