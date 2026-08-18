namespace AIE.Core.Models;

public class BoDonGia
{
    public int Id { get; set; }
    public string TenBo { get; set; } = string.Empty;
    public decimal GiaXang { get; set; }
    public decimal GiaDiezel { get; set; }
    public decimal GiaDien { get; set; }
    public string? GhiChu { get; set; }
    public string NgayTao { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
}
