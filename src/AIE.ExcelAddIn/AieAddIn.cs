using ExcelDna.Integration;

namespace AIE.ExcelAddIn;

/// <summary>
/// Lớp khởi tạo Add-in khi Excel tải.
/// </summary>
public class AieAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        // Code chạy khi Add-in được load vào Excel
        // Ví dụ: Đăng ký TaskPane, khởi tạo DB
    }

    public void AutoClose()
    {
        // Code chạy khi Add-in bị đóng
    }
}
