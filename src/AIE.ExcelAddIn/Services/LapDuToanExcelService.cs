using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using AIE.Core.Models;

namespace AIE.ExcelAddIn.Services;

/// <summary>
/// Service chuyên tương tác (đọc/ghi) với file Excel trong chế độ Lập Dự Toán.
/// </summary>
public class LapDuToanExcelService
{
    /// <summary>
    /// Đọc bảng tiên lượng (BOQ) từ Sheet hiện tại đang mở.
    /// Bắt đầu đọc từ dòng số 5 (Dưới header).
    /// </summary>
    public DuToan ReadBOQFromActiveSheet()
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        Worksheet ws = null;
        if (wb != null)
        {
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.StartsWith("DuToan_") || sheet.Name.StartsWith("DuToan "))
                {
                    ws = sheet;
                    break;
                }
            }
        }
        if (ws == null) ws = app.ActiveSheet as Worksheet;
        if (ws == null) throw new Exception("Không có Sheet nào đang mở.");

        var duToan = new DuToan();
        
        // Đọc thông tin công trình từ dòng 1 và 2 (Merge cell)
        string tenDuAn = GetCellValue(ws, 1, 1).Replace("TÊN DỰ ÁN:", "").Trim();
        string diaDiem = GetCellValue(ws, 2, 1).Replace("ĐỊA ĐIỂM:", "").Trim();
        
        duToan.TenCongTrinh = string.IsNullOrEmpty(tenDuAn) ? "Công trình mặc định" : tenDuAn;
        duToan.DiaDiem = string.IsNullOrEmpty(diaDiem) ? "Không xác định" : diaDiem;

        var hm = new HangMuc { STT = 1, TenHangMuc = "Hạng mục chung" };
        duToan.DanhSachHangMuc.Add(hm);

        Range usedRange = ws.UsedRange;
        int maxRow = usedRange.Rows.Count + usedRange.Row - 1;
        
        // Find the starting row by looking for 'STT' header
        int startRow = 5;
        for (int r = 1; r <= 10; r++)
        {
            if (GetCellValue(ws, r, 1) == "STT")
            {
                // STT could be merged with the row below it. The data starts after the header.
                // Let's assume data starts right after the row containing "Vật liệu" or just r + 2 if it's the new format
                string valBelow = GetCellValue(ws, r + 1, 1);
                if (string.IsNullOrWhiteSpace(valBelow)) 
                    startRow = r + 2; // Merged STT cell
                else
                    startRow = r + 1; // Single row header
                break;
            }
        }

        for (int r = startRow; r <= maxRow; r++)
        {
            string maHieu = GetCellValue(ws, r, 2); // Cột B
            if (string.IsNullOrWhiteSpace(maHieu)) continue;

            string ten = GetCellValue(ws, r, 3);    // Cột C
            string donVi = GetCellValue(ws, r, 4);  // Cột D
            string klStr = GetCellValue(ws, r, 5);  // Cột E
            string dgVLStr = GetCellValue(ws, r, 6);  // Cột F
            string dgNCStr = GetCellValue(ws, r, 7);  // Cột G
            string dgMayStr = GetCellValue(ws, r, 8); // Cột H

            decimal.TryParse(klStr, out decimal khoiLuong);
            decimal.TryParse(dgVLStr, out decimal donGiaVL);
            decimal.TryParse(dgNCStr, out decimal donGiaNC);
            decimal.TryParse(dgMayStr, out decimal donGiaMay);

            var dong = new DongDuToan
            {
                STT = r, // Dùng STT tạm bằng row để map ngược lại
                MaHieu = maHieu,
                TenCongTac = ten,
                DonVi = donVi,
                KhoiLuong = khoiLuong,
                DonGiaVL = donGiaVL,
                DonGiaNC = donGiaNC,
                DonGiaMay = donGiaMay
            };
            hm.DanhSachCongTac.Add(dong);
        }

        return duToan;
    }

    /// <summary>
    /// Gán đơn giá và công thức Thành tiền ngược lại các dòng tương ứng trên Excel.
    /// </summary>
    public void WriteDonGiaToExcel(DuToan duToan)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var ws = app.ActiveSheet as Worksheet;
        if (ws == null) return;

        foreach (var hm in duToan.DanhSachHangMuc)
        {
            foreach (var dong in hm.DanhSachCongTac)
            {
                int r = dong.STT; // Do lúc đọc ta lưu SoDongExcel vào STT
                if (r < 5) continue;

                // Gán Đơn Giá (Cột F, G, H)
                ws.Cells[r, 6].Value2 = dong.DonGiaVL;
                ws.Cells[r, 7].Value2 = dong.DonGiaNC;
                ws.Cells[r, 8].Value2 = dong.DonGiaMay;

                // Cột Thành tiền (I) = Khối lượng (E) * (ĐG VL + ĐG NC + ĐG Máy)
                // Đặt công thức: =E5*(F5+G5+H5)
                string formula = $"=E{r}*(F{r}+G{r}+H{r})";
                ws.Cells[r, 9].Formula = formula;
            }
        }
    }

    /// <summary>
    /// Ghi toàn bộ dữ liệu DuToan ra Sheet hiện tại (tái tạo bảng tiên lượng từ file .dt).
    /// </summary>
    public void WriteBOQToActiveSheet(DuToan duToan)
    {
        var app = (Application)ExcelDnaUtil.Application;
        
        // Nếu không có Workbook nào đang mở, tự động tạo mới
        if (app.Workbooks.Count == 0)
        {
            app.Workbooks.Add();
        }

        var ws = app.ActiveSheet as Worksheet;
        if (ws == null) throw new Exception("Không có Sheet nào đang mở.");

        app.ScreenUpdating = false;
        try
        {
            ws.Cells.Clear();

            // Set entire sheet font to Times New Roman, 12
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // Dòng 1: Tên dự án
            Range r1 = ws.Range["A1", "I1"];
            r1.Merge();
            r1.Value2 = ("TÊN DỰ ÁN: " + duToan.TenCongTrinh).ToUpper();
            r1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r1.Font.Bold = true;

            // Dòng 2: Địa điểm
            Range r2 = ws.Range["A2", "I2"];
            r2.Merge();
            r2.Value2 = ("ĐỊA ĐIỂM: " + duToan.DiaDiem).ToUpper();
            r2.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r2.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r2.Font.Bold = true;

            // Dòng 4 & 5: Header
            ws.Cells[4, 1] = "STT";
            ws.Cells[4, 2] = "Mã hiệu";
            ws.Cells[4, 3] = "Tên công tác";
            ws.Cells[4, 4] = "Đơn vị";
            ws.Cells[4, 5] = "Khối lượng";
            ws.Cells[4, 6] = "Đơn giá";
            ws.Cells[5, 6] = "Vật liệu";
            ws.Cells[5, 7] = "Nhân công";
            ws.Cells[5, 8] = "Máy thi công";
            ws.Cells[4, 9] = "Thành tiền";

            // Merge header cells
            ws.Range["A4:A5"].Merge();
            ws.Range["B4:B5"].Merge();
            ws.Range["C4:C5"].Merge();
            ws.Range["D4:D5"].Merge();
            ws.Range["E4:E5"].Merge();
            ws.Range["F4:H4"].Merge();
            ws.Range["I4:I5"].Merge();

            Range headerRange = ws.Range[ws.Cells[4, 1], ws.Cells[5, 9]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(200, 220, 240));
            headerRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            // Column widths
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[1]).ColumnWidth = 5;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[2]).ColumnWidth = 12;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[3]).ColumnWidth = 45;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[4]).ColumnWidth = 8;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[5]).ColumnWidth = 12;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[6]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[7]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[8]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns[9]).ColumnWidth = 18;

            if (ws.Name.StartsWith("Sheet"))
            {
                ws.Name = "DuToan_" + DateTime.Now.ToString("HHmmss");
            }

            // Dòng 5 trở đi: Dữ liệu
            int r = 6;
            int stt = 1;
            foreach (var hm in duToan.DanhSachHangMuc)
            {
                foreach (var dong in hm.DanhSachCongTac)
                {
                    dong.STT = r; // Cập nhật lại STT = dòng Excel để map ngược
                    ws.Cells[r, 1] = stt++;
                    ws.Cells[r, 2] = dong.MaHieu;
                    ws.Cells[r, 3] = dong.TenCongTac;
                    ws.Cells[r, 4] = dong.DonVi;
                    ws.Cells[r, 5] = (double)dong.KhoiLuong;

                    if (dong.DonGiaVL > 0 || dong.DonGiaNC > 0 || dong.DonGiaMay > 0)
                    {
                        ws.Cells[r, 6] = (double)dong.DonGiaVL;
                        ws.Cells[r, 7] = (double)dong.DonGiaNC;
                        ws.Cells[r, 8] = (double)dong.DonGiaMay;
                    }
                    
                    // Cột Thành tiền
                    ws.Cells[r, 9].Formula = $"=E{r}*(F{r}+G{r}+H{r})";

                    r++;
                }
            }

            // Định dạng số cho cột Khối lượng, đơn giá và Thành tiền
            if (r > 6)
            {
                Range dataRange = ws.Range[$"A6:I{r - 1}"];
                dataRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                dataRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                ws.Range[$"A6:A{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Range[$"C6:C{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                ws.Range[$"C6:C{r - 1}"].WrapText = true;
                ws.Range[$"D6:D{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                
                Range numberCols = ws.Range[$"E6:I{r - 1}"];
                numberCols.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                ws.Range[$"E6:E{r - 1}"].NumberFormat = "#,##0.000";
                ws.Range[$"F6:I{r - 1}"].NumberFormat = "#,##0";
            }
            
            // Freeze panes at row 5
            ws.Activate();
            app.ActiveWindow.FreezePanes = false;
            app.ActiveWindow.SplitRow = 5;
            app.ActiveWindow.SplitColumn = 0;
            app.ActiveWindow.FreezePanes = true;
        }
        finally
        {
            app.ScreenUpdating = true;
        }
    }

    private string GetCellValue(Worksheet ws, int row, int col)
    {
        try
        {
            Range range = ws.Cells[row, col];
            if (range.Value2 == null) return string.Empty;
            return range.Value2.ToString().Trim();
        }
        catch
        {
            return string.Empty;
        }
    }
}
