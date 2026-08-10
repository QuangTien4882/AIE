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

        int headerRow = config.DongBatDau - 1;
        if (headerRow < 1) headerRow = 1;

        int colTenTT38 = 0, colDviTT38 = 0, colDmTT38 = 0, colDmDiff = 0;

        // Xóa các cột liên quan đến Giá và Thành tiền ở sheet KQ_DinhMuc để đỡ rối
        if (ttIndex > 0) ((Range)resultSheet.Columns[ttIndex]).Delete();
        if (dgIndex > 0) ((Range)resultSheet.Columns[dgIndex]).Delete();

        // Chèn các cột TT38 sau cột Định mức dự toán
        int insertPos = dmIndex;
        if (insertPos > 0)
        {
            // Cần chèn 4 cột: Tên chuẩn, Đơn vị chuẩn, Định mức chuẩn, Chênh lệch
            for (int i = 0; i < 4; i++)
            {
                Range c = (Range)resultSheet.Columns[insertPos + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            
            colTenTT38 = insertPos + 1;
            colDviTT38 = insertPos + 2;
            colDmTT38 = insertPos + 3;
            colDmDiff = insertPos + 4;
        }

        // Determine rightmost inserted column to place "Ghi chú lỗi"
        int rightmostInserted = colDmDiff;
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
        var newCols = new List<int> { colTenTT38, colDviTT38, colDmTT38, colDmDiff, colGhiChu }.Where(x => x > 0).ToList();
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
                if (c == colTenTT38) headerRange.Value2 = "Tên Vật tư/NC/Máy (TT38)";
                if (c == colDviTT38) headerRange.Value2 = "Đơn vị (TT38)";
                if (c == colDmTT38) headerRange.Value2 = "Định mức (TT38)";
                if (c == colDmDiff) headerRange.Value2 = "Chênh lệch ĐM";
            }
            
            if (c == colTenTT38) resultSheet.Columns[c].ColumnWidth = 35;
            else if (c == colGhiChu) resultSheet.Columns[c].ColumnWidth = 35;
            else resultSheet.Columns[c].ColumnWidth = 15;

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
                        if (colTenTT38 > 0) resultSheet.Cells[rowExcel, colTenTT38].Value2 = saiLech.HaoPhiChuan.TenHaoPhi;
                        if (colDviTT38 > 0) resultSheet.Cells[rowExcel, colDviTT38].Value2 = saiLech.HaoPhiChuan.DonVi;
                        if (colDmTT38 > 0) resultSheet.Cells[rowExcel, colDmTT38].Value2 = saiLech.HaoPhiChuan.DinhMuc;
                        
                        if (colDmDiff > 0) 
                        {
                            resultSheet.Cells[rowExcel, colDmDiff].Value2 = saiLech.ChenhLechDinhMuc;
                            if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                                resultSheet.Cells[rowExcel, colDmDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }

                    if (saiLech.LoaiLoi != null && !saiLech.LoaiLoi.StartsWith("Thiếu hao phí"))
                    {
                        // Những lỗi trên dòng hao phí thì ghi chú tại đây
                        resultSheet.Cells[rowExcel, colGhiChu].Value2 = saiLech.LoaiLoi;
                        if (saiLech.LoaiLoi == "Khác tên gọi")
                            resultSheet.Cells[rowExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.DarkOrange);
                        else
                            resultSheet.Cells[rowExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.Red);
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

    public void ExportGiaVatTu(List<VatTuGiaModel> danhSachVatTu)
    {
        ExportGiaChoLoai(danhSachVatTu, AIE.Core.Enums.LoaiHaoPhi.VL, "KQ_GiaVL", "BẢNG THẨM ĐỊNH GIÁ VẬT LIỆU");
        ExportGiaChoLoai(danhSachVatTu, AIE.Core.Enums.LoaiHaoPhi.NC, "KQ_GiaNC", "BẢNG THẨM ĐỊNH GIÁ NHÂN CÔNG");
        ExportGiaChoLoai(danhSachVatTu, AIE.Core.Enums.LoaiHaoPhi.MAY, "KQ_GiaMay", "BẢNG THẨM ĐỊNH GIÁ MÁY THI CÔNG");
    }

    private void ExportGiaChoLoai(List<VatTuGiaModel> tatCaVatTu, AIE.Core.Enums.LoaiHaoPhi loai, string sheetName, string tieuDe)
    {
        var ds = tatCaVatTu.Where(x => x.LoaiHP == loai).ToList();
        if (ds.Count == 0) return;

        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        
        // Xóa sheet cũ nếu có
        foreach (Worksheet s in wb.Worksheets)
        {
            if (s.Name == sheetName)
            {
                app.DisplayAlerts = false;
                s.Delete();
                app.DisplayAlerts = true;
                break;
            }
        }

        // Tạo sheet mới
        var sheet = (Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
        sheet.Name = sheetName;

        // Định dạng tiêu đề
        sheet.Cells[1, 1].Value2 = tieuDe;
        sheet.Cells[1, 1].Font.Bold = true;
        sheet.Cells[1, 1].Font.Size = 14;
        
        // Header
        int row = 3;
        string[] headers = { "STT", "Mã hiệu TT38", "Tên Vật tư/NC/Máy", "Đơn vị", "Đơn giá Dự toán", "Đơn giá Chuẩn", "Chênh lệch" };
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = sheet.Cells[row, i + 1];
            cell.Value2 = headers[i];
            cell.Font.Bold = true;
            cell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // Data
        row++;
        for (int i = 0; i < ds.Count; i++)
        {
            var item = ds[i];
            sheet.Cells[row, 1].Value2 = i + 1;
            sheet.Cells[row, 2].Value2 = item.MaHieu;
            sheet.Cells[row, 3].Value2 = item.TenVatTu;
            sheet.Cells[row, 4].Value2 = item.DonVi;
            sheet.Cells[row, 5].Value2 = item.GiaDuToan;
            
            if (item.GiaChuan.HasValue)
            {
                sheet.Cells[row, 6].Value2 = item.GiaChuan.Value;
                sheet.Cells[row, 7].Value2 = item.ChenhLechGia;
                if (Math.Abs(item.ChenhLechGia) > 1)
                {
                    sheet.Cells[row, 7].Font.Color = ColorTranslator.ToOle(Color.Red);
                }
            }
            else
            {
                sheet.Cells[row, 6].Value2 = "Không có giá chuẩn";
                sheet.Cells[row, 6].Font.Color = ColorTranslator.ToOle(Color.DarkOrange);
            }

            Range dataRange = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 7]];
            dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            
            row++;
        }

        // Căn chỉnh độ rộng cột
        sheet.Columns[1].ColumnWidth = 5;
        sheet.Columns[2].ColumnWidth = 15;
        sheet.Columns[3].ColumnWidth = 45;
        sheet.Columns[4].ColumnWidth = 10;
        sheet.Columns[5].ColumnWidth = 18;
        sheet.Columns[6].ColumnWidth = 18;
        sheet.Columns[7].ColumnWidth = 18;
    }
}
