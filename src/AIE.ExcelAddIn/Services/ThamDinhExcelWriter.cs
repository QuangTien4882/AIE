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

        // Copy sheet xuống cuối cùng
        sourceSheet.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
        var resultSheet = (Worksheet)wb.ActiveSheet;
        resultSheet.Name = resultSheetName;

        int dmIndex = ColLetterToNumber(config.ColDinhMuc);
        int dgIndex = ColLetterToNumber(config.ColDonGia);
        int ttIndex = ColLetterToNumber(config.ColThanhTien);

        int currDm = dmIndex;

        int headerRow = config.DongBatDau - 1;
        if (headerRow < 1) headerRow = 1;

        // Chèn 1 dòng trống ngay dưới tiêu đề gốc để làm dòng phụ
        Range rowToInsert = (Range)resultSheet.Rows[headerRow + 1];
        rowToInsert.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

        // Cập nhật lại số dòng trong results do toàn bộ dữ liệu bị đẩy xuống 1 dòng
        foreach (var kq in results)
        {
            kq.DuToan.SoDongExcel++;
            foreach (var hp in kq.DuToan.DanhSachHaoPhi)
            {
                hp.SoDongExcel++;
            }
        }

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

        // Gộp các ô tiêu đề gốc theo chiều dọc (dòng headerRow và headerRow + 1)
        for (int c = 1; c <= lastCol; c++)
        {
            Range cellTop = resultSheet.Cells[headerRow, c];
            Range cellBottom = resultSheet.Cells[headerRow + 1, c];
            if (cellTop.Value2 != null || cellBottom.Value2 != null)
            {
                Range mergeRange = resultSheet.Range[cellTop, cellBottom];
                mergeRange.Merge();
                mergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
        }

        // Format new columns (headers and borders)
        var newCols = new List<int> { colTenTT38, colDviTT38, colDmTT38, colDmDiff, colGhiChu }.Where(x => x > 0).ToList();
        foreach (int c in newCols)
        {
            // Định dạng Header phụ (dòng headerRow + 1)
            Range subHeader = resultSheet.Cells[headerRow + 1, c];
            subHeader.Font.Bold = true;
            subHeader.Font.Color = ColorTranslator.ToOle(Color.Black);
            subHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            subHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            subHeader.WrapText = true;
            
            // Kế thừa font chữ từ ô bên cạnh
            if (c > 1)
            {
                Range prevCell = resultSheet.Cells[headerRow + 1, c - 1];
                subHeader.Font.Name = prevCell.Font.Name;
                subHeader.Font.Size = prevCell.Font.Size;
            }
            
            if (c == colGhiChu)
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                subHeader.Value2 = "Ghi chú lỗi (Cảnh báo)";
                
                // Ghi chú lỗi sẽ gộp cả 2 dòng
                Range mergeGhiChu = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[headerRow + 1, c]];
                mergeGhiChu.Merge();
                mergeGhiChu.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                // Ghi lại Text cho các cột
                if (c == colTenTT38) subHeader.Value2 = "Tên VT/NC/MTC";
                if (c == colDviTT38) subHeader.Value2 = "Đơn vị";
                if (c == colDmTT38) subHeader.Value2 = "Định mức";
                if (c == colDmDiff) subHeader.Value2 = "Chênh lệch ĐM";
            }
            
            if (c == colTenTT38) resultSheet.Columns[c].ColumnWidth = 35;
            else if (c == colGhiChu) resultSheet.Columns[c].ColumnWidth = 35;
            else resultSheet.Columns[c].ColumnWidth = 15;

            // Kẻ khung (Borders) cho toàn bộ dòng dữ liệu của cột mới
            if (lastRow >= headerRow + 1)
            {
                Range colRange = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[lastRow, c]];
                colRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
        }

        // Tạo tiêu đề gộp cho 4 cột TT38 (dòng headerRow)
        if (colTenTT38 > 0 && colDmDiff > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colTenTT38], resultSheet.Cells[headerRow, colDmDiff]];
            groupHeader.Merge();
            groupHeader.Value2 = "Định mức theo Thông tư 38";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
            
            // Kế thừa font chữ
            Range prevCell = resultSheet.Cells[headerRow, colTenTT38 - 1];
            groupHeader.Font.Name = prevCell.Font.Name;
            groupHeader.Font.Size = prevCell.Font.Size;
        }

        // Populate Data (Xử lý từ dưới lên trên để việc chèn dòng không làm sai lệch SoDongExcel của các dòng bên trên)
        foreach (var kq in results.OrderByDescending(x => x.DuToan.SoDongExcel))
        {
            var dt = kq.DuToan;
            
            // 1. Tô màu dòng công tác
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
            
            // In Tên công tác chuẩn TT38 để người dùng đối chiếu
            if (colTenTT38 > 0)
            {
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Value2 = kq.DinhMucChuan.TenCongTac;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Font.Bold = true;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Font.Name = resultSheet.Cells[dt.SoDongExcel, 3].Font.Name;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204)); // Vàng nhạt để nổi bật
            }

            // 2. Chèn dòng cho các "Hao phí thiếu" (Có trong TT38 nhưng không có trong Dự toán)
            var thieuItems = kq.DanhSachSaiLech.Where(x => x.HaoPhiDuToan == null && x.HaoPhiChuan != null).ToList();
            var loaiHps = new[] { AIE.Core.Enums.LoaiHaoPhi.MAY, AIE.Core.Enums.LoaiHaoPhi.NC, AIE.Core.Enums.LoaiHaoPhi.VL };
            
            foreach (var loai in loaiHps)
            {
                var thieuLoai = thieuItems.Where(x => x.HaoPhiChuan.LoaiHaoPhi == loai).ToList();
                if (thieuLoai.Any())
                {
                    // Tìm dòng thích hợp để chèn (ngay dưới hao phí cuối cùng cùng loại, hoặc loại trước đó)
                    int insertRow = dt.SoDongExcel;
                    var sameLoai = dt.DanhSachHaoPhi.Where(x => x.Loai == loai).ToList();
                    if (sameLoai.Any()) insertRow = sameLoai.Max(x => x.SoDongExcel);
                    else if (loai == AIE.Core.Enums.LoaiHaoPhi.MAY)
                    {
                        var nc = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.NC).ToList();
                        if (nc.Any()) insertRow = nc.Max(x => x.SoDongExcel);
                        else
                        {
                            var vl = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.VL).ToList();
                            if (vl.Any()) insertRow = vl.Max(x => x.SoDongExcel);
                        }
                    }
                    else if (loai == AIE.Core.Enums.LoaiHaoPhi.NC)
                    {
                        var vl = dt.DanhSachHaoPhi.Where(x => x.Loai == AIE.Core.Enums.LoaiHaoPhi.VL).ToList();
                        if (vl.Any()) insertRow = vl.Max(x => x.SoDongExcel);
                    }

                    // Chèn N dòng
                    for (int i = 0; i < thieuLoai.Count; i++)
                    {
                        int newRow = insertRow + 1 + i;
                        Range insertedRow = (Range)resultSheet.Rows[newRow];
                        insertedRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        
                        // Clear định dạng có thể bị kế thừa từ dòng công tác
                        Range newRowRange = resultSheet.Range[resultSheet.Cells[newRow, 1], resultSheet.Cells[newRow, lastCol]];
                        newRowRange.Interior.ColorIndex = 0;
                        newRowRange.Font.Bold = false;
                        
                        var sl = thieuLoai[i];
                        if (colTenTT38 > 0) 
                        {
                            resultSheet.Cells[newRow, colTenTT38].Value2 = sl.HaoPhiChuan.TenHaoPhi;
                            resultSheet.Cells[newRow, colTenTT38].Font.Name = resultSheet.Cells[insertRow, 3].Font.Name;
                            resultSheet.Cells[newRow, colTenTT38].Font.Size = resultSheet.Cells[insertRow, 3].Font.Size;
                        }
                        if (colDviTT38 > 0) 
                        {
                            resultSheet.Cells[newRow, colDviTT38].Value2 = sl.HaoPhiChuan.DonVi;
                            resultSheet.Cells[newRow, colDviTT38].Font.Name = resultSheet.Cells[insertRow, 3].Font.Name;
                            resultSheet.Cells[newRow, colDviTT38].Font.Size = resultSheet.Cells[insertRow, 3].Font.Size;
                        }
                        if (colDmTT38 > 0) 
                        {
                            resultSheet.Cells[newRow, colDmTT38].Value2 = sl.HaoPhiChuan.DinhMuc;
                            resultSheet.Cells[newRow, colDmTT38].Font.Name = resultSheet.Cells[insertRow, 3].Font.Name;
                            resultSheet.Cells[newRow, colDmTT38].Font.Size = resultSheet.Cells[insertRow, 3].Font.Size;
                        }
                        
                        resultSheet.Cells[newRow, colGhiChu].Value2 = $"[Thiếu hao phí] TT38 có '{sl.HaoPhiChuan.TenHaoPhi}'";
                        resultSheet.Cells[newRow, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.Red);
                        resultSheet.Cells[newRow, colGhiChu].Font.Name = resultSheet.Cells[insertRow, 3].Font.Name;
                        resultSheet.Cells[newRow, colGhiChu].Font.Size = resultSheet.Cells[insertRow, 3].Font.Size;

                        // Cập nhật lại SoDongExcel cho các hao phí dự toán bị đẩy xuống
                        foreach (var hp in dt.DanhSachHaoPhi)
                        {
                            if (hp.SoDongExcel >= newRow) hp.SoDongExcel++;
                        }
                    }
                }
            }

            // 3. Ghi dữ liệu cho các hao phí Đã Match hoặc Thừa
            foreach (var hpDuToan in dt.DanhSachHaoPhi)
            {
                var rowExcel = hpDuToan.SoDongExcel;
                var saiLech = kq.DanhSachSaiLech.FirstOrDefault(x => x.HaoPhiDuToan == hpDuToan);
                
                // Lấy font chuẩn từ cột dự toán gốc để đồng bộ
                string fontName = resultSheet.Cells[rowExcel, 3].Font.Name?.ToString() ?? "Times New Roman";
                double fontSize = resultSheet.Cells[rowExcel, 3].Font.Size;
                
                if (saiLech != null)
                {
                    if (saiLech.HaoPhiChuan == null)
                    {
                        // Explicitly label as Thừa hao phí
                        resultSheet.Cells[rowExcel, colGhiChu].Value2 = "Hao phí thừa";
                        resultSheet.Cells[rowExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.Red);
                        resultSheet.Cells[rowExcel, colGhiChu].Font.Bold = true;
                        resultSheet.Cells[rowExcel, colGhiChu].Font.Name = fontName;
                        resultSheet.Cells[rowExcel, colGhiChu].Font.Size = fontSize;
                    }
                    else
                    {
                        if (colTenTT38 > 0) 
                        {
                            resultSheet.Cells[rowExcel, colTenTT38].Value2 = saiLech.HaoPhiChuan.TenHaoPhi;
                            resultSheet.Cells[rowExcel, colTenTT38].Font.Name = fontName;
                            resultSheet.Cells[rowExcel, colTenTT38].Font.Size = fontSize;
                        }
                        if (colDviTT38 > 0) 
                        {
                            resultSheet.Cells[rowExcel, colDviTT38].Value2 = saiLech.HaoPhiChuan.DonVi;
                            resultSheet.Cells[rowExcel, colDviTT38].Font.Name = fontName;
                            resultSheet.Cells[rowExcel, colDviTT38].Font.Size = fontSize;
                        }
                        if (colDmTT38 > 0) 
                        {
                            resultSheet.Cells[rowExcel, colDmTT38].Value2 = saiLech.HaoPhiChuan.DinhMuc;
                            resultSheet.Cells[rowExcel, colDmTT38].Font.Name = fontName;
                            resultSheet.Cells[rowExcel, colDmTT38].Font.Size = fontSize;
                        }
                        
                        if (colDmDiff > 0) 
                        {
                            resultSheet.Cells[rowExcel, colDmDiff].Value2 = saiLech.ChenhLechDinhMuc;
                            resultSheet.Cells[rowExcel, colDmDiff].Font.Name = fontName;
                            resultSheet.Cells[rowExcel, colDmDiff].Font.Size = fontSize;
                            if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                                resultSheet.Cells[rowExcel, colDmDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        if (saiLech.LoaiLoi != null && !saiLech.LoaiLoi.StartsWith("Thiếu hao phí"))
                        {
                            resultSheet.Cells[rowExcel, colGhiChu].Value2 = saiLech.LoaiLoi;
                            resultSheet.Cells[rowExcel, colGhiChu].Font.Name = fontName;
                            resultSheet.Cells[rowExcel, colGhiChu].Font.Size = fontSize;
                            if (saiLech.LoaiLoi == "Khác tên gọi")
                                resultSheet.Cells[rowExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.DarkOrange);
                            else
                                resultSheet.Cells[rowExcel, colGhiChu].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }
                }
            }
        }

        // Cố định dòng tiêu đề (Freeze Panes)
        try
        {
            resultSheet.Activate();
            
            // Cuộn màn hình lên góc trái trên cùng
            resultSheet.Application.ActiveWindow.ScrollRow = 1;
            resultSheet.Application.ActiveWindow.ScrollColumn = 1;
            
            // Cố định dưới dòng headerRow + 1 (vì tiêu đề hiện tại có 2 dòng)
            resultSheet.Application.ActiveWindow.SplitRow = headerRow + 1;
            resultSheet.Application.ActiveWindow.SplitColumn = 0;
            resultSheet.Application.ActiveWindow.FreezePanes = true;

            // Thiết lập Wrap Text cho toàn bộ vùng dữ liệu (từ dòng headerRow + 2 trở xuống)
            if (lastRow > headerRow + 1)
            {
                Range dataRange = resultSheet.Range[resultSheet.Cells[headerRow + 2, 1], resultSheet.Cells[lastRow, colGhiChu]];
                dataRange.WrapText = true;
                dataRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
        }
        catch { }
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
