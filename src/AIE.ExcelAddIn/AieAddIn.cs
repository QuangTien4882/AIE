using ExcelDna.Integration;

namespace AIE.ExcelAddIn;

/// <summary>
/// Lớp khởi tạo Add-in khi Excel tải.
/// </summary>
public class AieAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        // Khởi tạo DB và seed dữ liệu nhân công nếu chưa có
        try
        {
            var db = new AIE.Data.DatabaseManager();
            db.SeedNhanCong();
        }
        catch { /* Bỏ qua lỗi seed */ }
    }

    public void AutoClose()
    {
        // Code chạy khi Add-in bị đóng
    }
}
