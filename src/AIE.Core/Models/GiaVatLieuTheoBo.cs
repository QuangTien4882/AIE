namespace AIE.Core.Models;

public class GiaVatLieuTheoBo
{
    public int Id { get; set; }
    public int BoDonGiaId { get; set; }
    public string MaVL { get; set; } = string.Empty;
    public string? NguonCungCap { get; set; }
    public decimal GiaGoc { get; set; }
    
    // Cước vận chuyển cơ bản
    public decimal CuLy_Km { get; set; }
    public string? LoaiDuong { get; set; }
    public decimal CuocVC { get; set; }
    
    // Tổng giá
    public decimal GiaHienTruong { get; set; }
    
    // Chi tiết cước vận chuyển & bốc xếp
    public decimal ChiPhiBocXep { get; set; }
    public decimal CuocVCOTo { get; set; }
    public decimal CuocVCBo { get; set; }

    /// <summary>
    /// 1 = Nguồn có giá thấp nhất được chọn để tính toán
    /// 0 = Các nguồn khác
    /// </summary>
    public int DuocChon { get; set; }
}
