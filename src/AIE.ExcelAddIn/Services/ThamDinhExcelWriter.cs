using AIE.Core.Models;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace AIE.ExcelAddIn.Services;

public class ThamDinhExcelWriter
{
    public void ExportResult(ThamDinhConfig config, List<KetQuaCongTacThamDinh> results)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        
        // Find source sheet
        Worksheet sourceSheet = null;
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name == config.SheetName)
            {
                sourceSheet = sheet;
                break;
            }
        }
        
        if (sourceSheet == null) throw new Exception("Không tìm thấy sheet nguồn.");

        // Delete old result sheet if exists
        string resultSheetName = "KQ_ThamDinh";
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name == resultSheetName)
            {
                app.DisplayAlerts = false;
                sheet.Delete();
                app.DisplayAlerts = true;
                break;
            }
        }

        // Copy sheet
        sourceSheet.Copy(After: sourceSheet);
        var resultSheet = (Worksheet)wb.ActiveSheet;
        resultSheet.Name = resultSheetName;

        int dmIndex = ColLetterToNumber(config.ColDinhMuc);
        int dgIndex = ColLetterToNumber(config.ColDonGia);
        int ttIndex = ColLetterToNumber(config.ColThanhTien);

        int currDm = dmIndex;
        int currDg = dgIndex;
        int currTt = ttIndex;

        int colDmTT38 = 0, colDmDiff = 0, colDgTT38 = 0, colDgDiff = 0, colTtTT38 = 0, colTtDiff = 0;

        // Insert from right to left to avoid index shifting problems
        var insertPoints = new List<int> { dmIndex, dgIndex, ttIndex }.Where(x => x > 0).Distinct().OrderByDescending(x => x).ToList();
        
        int headerRow = config.DongBatDau - 1;
        if (headerRow < 1) headerRow = 1;

        foreach (int p in insertPoints)
        {
            // Insert 2 columns right after p
            Range c1 = (Range)resultSheet.Columns[p + 1];
            c1.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            Range c2 = (Range)resultSheet.Columns[p + 2];
            c2.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

            // Shift indices if they are > p
            if (currDm > p) currDm += 2;
            if (colDmTT38 > p) colDmTT38 += 2;
            if (colDmDiff > p) colDmDiff += 2;
            
            if (currDg > p) currDg += 2;
            if (colDgTT38 > p) colDgTT38 += 2;
            if (colDgDiff > p) colDgDiff += 2;
            
            if (currTt > p) currTt += 2;
            if (colTtTT38 > p) colTtTT38 += 2;
            if (colTtDiff > p) colTtDiff += 2;

            if (p == ttIndex)
            {
                colTtTT38 = p + 1;
                colTtDiff = p + 2;
            }
            if (p == dgIndex)
            {
                colDgTT38 = p + 1;
                colDgDiff = p + 2;
            }
            if (p == dmIndex)
            {
                colDmTT38 = p + 1;
                colDmDiff = p + 2;
            }
        }

        // Determine rightmost inserted column to place "Ghi chú lỗi"
        int rightmostInserted = Math.Max(colDmDiff, Math.Max(colDgDiff, colTtDiff));
        int colGhiChu = rightmostInserted + 1;
        if (rightmostInserted > 0)
        {
            Range cGhiChu = (Range)resultSheet.Columns[colGhiChu];
            cGhiChu.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }
        else
        {
            Range usedRange = resultSheet.UsedRange;
            colGhiChu = usedRange.Columns.Count + usedRange.Column;
        }

        // Find last column for background coloring
        Range lastColRange = resultSheet.UsedRange;
        int lastCol = lastColRange.Columns.Count + lastColRange.Column - 1;

        int lastRow = headerRow;
        if (results.Count > 0)
        {
            lastRow = results.Max(x => 
                Math.Max(x.DuToan.SoDongExcel, 
                x.DuToan.DanhSachHaoPhi.Count > 0 ? x.DuToan.DanhSachHaoPhi.Max(h => h.SoDongExcel) : 0)
            );
        }

        // Format new columns (headers and borders)
        var newCols = new List<int> { colDmTT38, colDmDiff, colDgTT38, colDgDiff, colTtTT38, colTtDiff, colGhiChu }.Where(x => x > 0).ToList();
        foreach (int c in newCols)
        {
            // Định dạng Header
            Range headerRange = resultSheet.Cells[headerRow, c];
            headerRange.UnMerge(); // Fix lỗi bị che khuất text do merge cell từ cột cũ
            headerRange.Font.Name = "Arial"; // Bắt buộc dùng Unicode font để chống lỗi font VNI
            headerRange.Font.Bold = true;
            headerRange.Font.Color = ColorTranslator.ToOle(Color.Black);
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.WrapText = true;
            
            if (c == colGhiChu)
            {
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                headerRange.Value2 = "Ghi chú lỗi (Cảnh báo)";
            }
            else
            {
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                
                // Ghi lại Text cho các cột
                if (c == colDmTT38) headerRange.Value2 = "Định mức TT38";
                if (c == colDmDiff) headerRange.Value2 = "Chênh lệch ĐM";
                if (c == colDgTT38) headerRange.Value2 = "Đơn giá thẩm định";
                if (c == colDgDiff) headerRange.Value2 = "Chênh lệch ĐG";
                if (c == colTtTT38) headerRange.Value2 = "Thành tiền thẩm định";
                if (c == colTtDiff) headerRange.Value2 = "Chênh lệch TT";
            }
            
            resultSheet.Columns[c].ColumnWidth = (c == colGhiChu) ? 35 : 18;

            // Kẻ khung (Borders) và Font cho toàn bộ dòng dữ liệu của cột mới
            if (lastRow >= headerRow)
            {
                Range colRange = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[lastRow, c]];
                colRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                colRange.Font.Name = "Arial";
            }
        }

        // Populate Data
        foreach (var kq in results)
        {
            var dt = kq.DuToan;
            
            // Tô màu dòng công tác
            Range ctRange = resultSheet.Range[resultSheet.Cells[dt.SoDongExcel, 1], resultSheet.Cells[dt.SoDongExcel, lastCol]];
            
            if (kq.DinhMucChuan == null)
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 230, 153)); // Cam nhạt
                resultSheet.Cells[dt.SoDongExcel, colGhiChu].Value2 = "Mã hiệu không tồn tại trong TT38";
                continue;
            }

            if (kq.DaDat)
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(226, 239, 218)); // Xanh nhạt
            }
            else
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(252, 228, 214)); // Đỏ nhạt
            }

            // Ghi chú hao phí thừa/thiếu vào dòng công tác
            var loiThieu = kq.DanhSachSaiLech.Where(x => x.LoaiLoi == "Thiếu hao phí" || x.LoaiLoi == "Hao phí thừa / Không khớp").ToList();
            if (loiThieu.Any())
            {
                string loiStr = string.Join("\n", loiThieu.Select(x => $"[{x.LoaiLoi}] {x.MoTa}"));
                resultSheet.Cells[dt.SoDongExcel, colGhiChu].Value2 = loiStr;
                resultSheet.Cells[dt.SoDongExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.Red);
            }

            // Fill row-by-row
            foreach (var hpDuToan in dt.DanhSachHaoPhi)
            {
                var rowExcel = hpDuToan.SoDongExcel;
                var saiLech = kq.DanhSachSaiLech.FirstOrDefault(x => x.SoDongExcel == rowExcel);
                
                if (saiLech != null)
                {
                    if (saiLech.HaoPhiChuan != null)
                    {
                        if (colDmTT38 > 0) resultSheet.Cells[rowExcel, colDmTT38].Value2 = saiLech.HaoPhiChuan.DinhMuc;
                        if (colDmDiff > 0) 
                        {
                            resultSheet.Cells[rowExcel, colDmDiff].Value2 = saiLech.ChenhLechDinhMuc;
                            if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                                resultSheet.Cells[rowExcel, colDmDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }

                    if (saiLech.DonGiaChuan.HasValue)
                    {
                        if (colDgTT38 > 0) resultSheet.Cells[rowExcel, colDgTT38].Value2 = saiLech.DonGiaChuan.Value;
                        if (colDgDiff > 0)
                        {
                            resultSheet.Cells[rowExcel, colDgDiff].Value2 = saiLech.ChenhLechDonGia;
                            if (Math.Abs(saiLech.ChenhLechDonGia) > 1)
                                resultSheet.Cells[rowExcel, colDgDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                        
                        if (colTtTT38 > 0 && saiLech.HaoPhiChuan != null && colDmTT38 > 0 && colDgTT38 > 0) 
                        {
                            resultSheet.Cells[rowExcel, colTtTT38].Formula = $"={GetColumnLetter(colDmTT38)}{rowExcel}*{GetColumnLetter(colDgTT38)}{rowExcel}";
                        }
                        
                        if (colTtDiff > 0 && colTtTT38 > 0 && currTt > 0)
                        {
                            resultSheet.Cells[rowExcel, colTtDiff].Formula = $"={GetColumnLetter(colTtTT38)}{rowExcel}-{GetColumnLetter(currTt)}{rowExcel}";
                        }
                    }
                    
                    if (saiLech.LoaiLoi != null && !saiLech.LoaiLoi.StartsWith("Thiếu hao phí"))
                    {
                        // Những lỗi trên dòng hao phí (như sai định mức, sai đơn vị) thì ghi chú tại đây
                        resultSheet.Cells[rowExcel, colGhiChu].Value2 = saiLech.LoaiLoi;
                    }
                }
            }
        }
    }

    private int ColLetterToNumber(string letter)
    {
        if (string.IsNullOrEmpty(letter)) return 0;
        int col = 0;
        foreach (char c in letter.ToUpper())
        {
            col = col * 26 + (c - 'A' + 1);
        }
        return col;
    }

    private string GetColumnLetter(int colIndex)
    {
        if (colIndex <= 0) return "";
        int div = colIndex;
        string colLetter = string.Empty;
        int mod = 0;

        while (div > 0)
        {
            mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = (int)((div - mod) / 26);
        }
        return colLetter;
    }
}
