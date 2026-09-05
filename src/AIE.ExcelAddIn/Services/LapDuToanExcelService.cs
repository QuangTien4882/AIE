using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using AIE.Core.Models;
using AIE.ExcelAddIn.Helpers;

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
        
        // Đọc thông tin công trình từ dòng 1, 2, 3 (Merge cell)
        string r1Val = GetCellValue(ws, 1, 1).Trim();
        string r2Val = GetCellValue(ws, 2, 1).Trim();
        string r3Val = GetCellValue(ws, 3, 1).Trim();

        string tenDuAn = "";
        string diaDiem = "";

        if (r1Val.ToUpper().Contains("BẢNG DỰ TOÁN"))
        {
            tenDuAn = r2Val.Replace("TÊN DỰ ÁN:", "").Replace("Dự án:", "").Trim();
            diaDiem = r3Val.Replace("ĐỊA ĐIỂM XÂY DỰNG:", "").Replace("Địa điểm xây dựng:", "").Replace("ĐỊA ĐIỂM:", "").Replace("Địa điểm:", "").Trim();
        }
        else
        {
            tenDuAn = r1Val.Replace("TÊN DỰ ÁN:", "").Replace("Dự án:", "").Trim();
            diaDiem = r2Val.Replace("ĐỊA ĐIỂM XÂY DỰNG:", "").Replace("Địa điểm xây dựng:", "").Replace("ĐỊA ĐIỂM:", "").Replace("Địa điểm:", "").Trim();
        }
        
        duToan.TenCongTrinh = string.IsNullOrEmpty(tenDuAn) ? "Công trình mặc định" : tenDuAn;
        duToan.DiaDiem = string.IsNullOrEmpty(diaDiem) ? "Không xác định" : diaDiem;

        // Đọc Vùng áp dụng từ ô J1 (lưu dưới dạng int enum)
        string vungStr = GetCellValue(ws, 1, 10);
        if (int.TryParse(vungStr, out int vungInt) && System.Enum.IsDefined(typeof(AIE.Core.Enums.Vung), vungInt))
        {
            duToan.VungApDung = (AIE.Core.Enums.Vung)vungInt;
        }

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
                // STT is merged with the row below it. Data starts after 2 header rows.
                startRow = r + 2;
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
                if (r < 4) continue;

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

            // Dòng 1: Tiêu đề sheet
            Range r0 = ws.Range["A1", "I1"];
            r0.Merge();
            r0.Value2 = "BẢNG DỰ TOÁN CHI TIẾT";
            r0.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r0.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r0.Font.Bold = true;
            r0.Font.Size = 14;

            // Dòng 2: Dự án (đồng bộ với TongMucDauTu/TH_DuToan)
            string tenDuAn = UIHelper.ChuanHoaChuThuong(duToan.TenCongTrinh);
            if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
            Range r1 = ws.Range["A2", "I2"];
            r1.Merge();
            r1.Value2 = "Dự án: " + tenDuAn;
            r1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r1.Font.Bold = true;
            r1.Font.Size = 12;

            // Dòng 3: Địa điểm xây dựng (đồng bộ với TongMucDauTu/TH_DuToan)
            string diaDiem = UIHelper.ChuanHoaChuThuong(duToan.DiaDiem);
            if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";
            Range r2 = ws.Range["A3", "I3"];
            r2.Merge();
            r2.Value2 = "Địa điểm xây dựng: " + diaDiem;
            r2.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            r2.VerticalAlignment = XlVAlign.xlVAlignCenter;
            r2.Font.Bold = true;
            r2.Font.Size = 12;

            // Dòng 4 & 5: Header 2 dòng chuẩn
            ws.Cells[4, 1] = "STT";
            ws.Cells[4, 2] = "Mã hiệu";
            ws.Cells[4, 3] = "Tên công tác";
            ws.Cells[4, 4] = "Đơn vị";
            ws.Cells[4, 5] = "Khối lượng";
            ws.Cells[4, 6] = "Đơn giá";
            ws.Cells[4, 9] = "Thành tiền";

            ws.Cells[5, 6] = "Vật liệu";
            ws.Cells[5, 7] = "Nhân công";
            ws.Cells[5, 8] = "Máy thi công";

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

            // Dòng 6 trở đi: Dữ liệu
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
            if (r > 5)
            {
                Range dataRange = ws.Range[$"A5:I{r - 1}"];
                dataRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                dataRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                ws.Range[$"A5:A{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Range[$"C5:C{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                ws.Range[$"C5:C{r - 1}"].WrapText = true;
                ws.Range[$"D5:D{r - 1}"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                
                Range numberCols = ws.Range[$"E5:I{r - 1}"];
                numberCols.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                ws.Range[$"E5:E{r - 1}"].NumberFormat = "#,##0.000";
                ws.Range[$"F5:I{r - 1}"].NumberFormat = "#,##0";
            }
            
            // Freeze panes at row 4
            ws.Activate();
            app.ActiveWindow.FreezePanes = false;
            app.ActiveWindow.SplitRow = 4;
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
