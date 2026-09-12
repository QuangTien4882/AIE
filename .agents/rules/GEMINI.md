# AIE Project Rules

## Quy tắc chung
- Dự án sử dụng .NET Framework 4.8 (không phải .NET Core/6/7/8), target `net48`
- Ngôn ngữ C# với `LangVersion: latest`
- UI sử dụng WinForms (System.Windows.Forms), KHÔNG sử dụng WPF
- Excel integration qua ExcelDna + Microsoft.Office.Interop.Excel
- Database là SQLite + Dapper (không dùng Entity Framework)

## Convention
- Đặt tên biến, class, method theo tiếng Việt không dấu (VD: DuToan, ChiPhiXayDung, TinhGiaHienTruong)
- Comment bằng tiếng Việt có dấu
- Mỗi form WinForms là 1 file .cs duy nhất (code-behind, không dùng Designer)
- Tất cả ribbon handlers phải có try-catch bao bọc, hiển thị lỗi qua MessageBox

## Kiến trúc
- `AieRibbon.CurrentDuToan` là object DuToan đang làm việc (in-memory, static)
- Dữ liệu được serialize/deserialize qua file `.dt` (JSON format, Newtonsoft.Json)
- Giá VL/NC/Máy lấy từ Bộ Đơn Giá (BoDonGia) trong SQLite
- Mọi tính toán chi phí tuân thủ công thức Thông tư 36/2026/TT-BXD
- Định mức tỷ lệ % QLDA/Tư vấn tra bảng nội suy theo Thông tư 38/2026/TT-BXD
- Làm tròn theo RoundingRules: đơn giá/thành tiền = 0 chữ số thập phân, khối lượng = 3, hao phí = 5

## Build
- Solution file: `AIE.slnx`
- Build command: `dotnet build AIE.slnx --configuration Debug`
- PowerShell trên Windows: dùng `;` thay cho `&&` để nối lệnh
