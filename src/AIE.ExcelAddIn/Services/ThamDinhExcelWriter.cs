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

        // Delete old result sheets if exist
        string[] sheetsToDelete = { "KQ_ThamDinh", "KQ_GiaVL", "KQ_GiaNC", "KQ_GiaMay" };
        foreach (string sheetName in sheetsToDelete)
        {
            foreach (Worksheet sheet in wb.Worksheets)
            {
                if (sheet.Name == sheetName)
                {
                    app.DisplayAlerts = false;
                    sheet.Delete();
                    app.DisplayAlerts = true;
                    break;
                }
            }
        }

        // Copy sheet xuống cuối cùng
        sourceSheet.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
        var resultSheet = (Worksheet)wb.ActiveSheet;
        resultSheet.Name = "KQ_ThamDinh";

        int dmIndex = ColLetterToNumber(config.ColDinhMuc);
        int dgIndex = ColLetterToNumber(config.ColDonGia);
        int ttIndex = ColLetterToNumber(config.ColThanhTien);

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

        // ============================================================
        // CHÈN CÁC CỘT MỚI (KHÔNG xóa cột Đơn giá và Thành tiền)
        // ============================================================

        // --- Bước 1: Chèn 4 cột TT38 sau cột Định mức ---
        int colTenTT38 = 0, colDviTT38 = 0, colDmTT38 = 0, colDmDiff = 0;
        if (dmIndex > 0)
        {
            for (int i = 0; i < 4; i++)
            {
                Range c = (Range)resultSheet.Columns[dmIndex + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colTenTT38 = dmIndex + 1;
            colDviTT38 = dmIndex + 2;
            colDmTT38 = dmIndex + 3;
            colDmDiff = dmIndex + 4;
        }

        // Cập nhật vị trí Đơn giá và Thành tiền (dịch phải 4 cột do chèn TT38)
        int colDonGiaDuToan = dgIndex > 0 ? dgIndex + 4 : 0;
        int colThanhTienDuToan = ttIndex > 0 ? ttIndex + 4 : 0;

        // --- Bước 2: Chèn 2 cột sau Đơn giá dự toán (ĐG chuẩn, Chênh lệch ĐG) ---
        int colDonGiaChuan = 0, colChenhLechDG = 0;
        if (colDonGiaDuToan > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[colDonGiaDuToan + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colDonGiaChuan = colDonGiaDuToan + 1;
            colChenhLechDG = colDonGiaDuToan + 2;
            // Thành tiền bị dịch thêm 2
            if (colThanhTienDuToan > 0) colThanhTienDuToan += 2;
        }

        // --- Bước 3: Chèn 2 cột sau Thành tiền dự toán (TT thẩm định, Chênh lệch TT) ---
        int colThanhTienTD = 0, colChenhLechTT = 0;
        if (colThanhTienDuToan > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[colThanhTienDuToan + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colThanhTienTD = colThanhTienDuToan + 1;
            colChenhLechTT = colThanhTienDuToan + 2;
        }

        // --- Bước 4: Chèn 2 cột Ghi chú lỗi ở cuối ---
        int rightmost = new[] { colChenhLechTT, colChenhLechDG, colDmDiff }.Where(x => x > 0).DefaultIfEmpty(0).Max();
        int colGhiChuDM, colGhiChuDG;
        if (rightmost > 0)
        {
            for (int i = 0; i < 2; i++)
            {
                Range c = (Range)resultSheet.Columns[rightmost + 1];
                c.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            colGhiChuDM = rightmost + 1;
            colGhiChuDG = rightmost + 2;
        }
        else
        {
            Range usedRange = resultSheet.UsedRange;
            colGhiChuDM = usedRange.Columns.Count + usedRange.Column;
            colGhiChuDG = colGhiChuDM + 1;
        }

        // ============================================================
        // TÌM VÙNG DỮ LIỆU
        // ============================================================
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

        // ============================================================
        // THIẾT LẬP FONT SIZE 12 CHO TOÀN BỘ SHEET
        // ============================================================
        resultSheet.Cells.Font.Size = 12;

        // ============================================================
        // ĐỊNH DẠNG TIÊU ĐỀ
        // ============================================================

        // Danh sách các cột mới (chèn thêm) - dùng để phân biệt với cột gốc
        var newCols = new List<int> { colTenTT38, colDviTT38, colDmTT38, colDmDiff,
            colDonGiaChuan, colChenhLechDG, colThanhTienTD, colChenhLechTT, colGhiChuDM, colGhiChuDG }
            .Where(x => x > 0).ToHashSet();

        // Gộp các ô tiêu đề GỐC theo chiều dọc (dòng headerRow và headerRow + 1)
        for (int c = 1; c <= lastCol; c++)
        {
            if (newCols.Contains(c)) continue; // Bỏ qua cột mới

            Range cellTop = resultSheet.Cells[headerRow, c];
            Range cellBottom = resultSheet.Cells[headerRow + 1, c];
            if (cellTop.Value2 != null || cellBottom.Value2 != null)
            {
                Range mergeRange = resultSheet.Range[cellTop, cellBottom];
                mergeRange.Merge();
                mergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
        }

        // Lấy font name chuẩn từ ô tiêu đề gốc
        string baseFontName = resultSheet.Cells[headerRow, 1].Font.Name?.ToString() ?? "Times New Roman";

        // Định dạng các cột mới (sub-header ở dòng headerRow + 1)
        foreach (int c in newCols)
        {
            Range subHeader = resultSheet.Cells[headerRow + 1, c];
            subHeader.Font.Bold = true;
            subHeader.Font.Size = 12;
            subHeader.Font.Name = baseFontName;
            subHeader.Font.Color = ColorTranslator.ToOle(Color.Black);
            subHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            subHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            subHeader.WrapText = true;
            
            if (c == colGhiChuDM)
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.Orange);
                subHeader.Value2 = "Lỗi Định mức";
                // Ghi chú gộp cả 2 dòng header
                Range mergeGhiChu = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[headerRow + 1, c]];
                mergeGhiChu.Merge();
                mergeGhiChu.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else if (c == colGhiChuDG)
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 192, 0));
                subHeader.Value2 = "Lỗi Đơn giá";
                // Ghi chú gộp cả 2 dòng header
                Range mergeGhiChu = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[headerRow + 1, c]];
                mergeGhiChu.Merge();
                mergeGhiChu.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else
            {
                subHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                if (c == colTenTT38) subHeader.Value2 = "Tên VT/NC/MTC";
                if (c == colDviTT38) subHeader.Value2 = "Đơn vị";
                if (c == colDmTT38) subHeader.Value2 = "Định mức";
                if (c == colDmDiff) subHeader.Value2 = "Chênh lệch ĐM";
                if (c == colDonGiaChuan) subHeader.Value2 = "Đơn giá chuẩn";
                if (c == colChenhLechDG) subHeader.Value2 = "Chênh lệch ĐG";
                if (c == colThanhTienTD) subHeader.Value2 = "Thành tiền (TĐ)";
                if (c == colChenhLechTT) subHeader.Value2 = "Chênh lệch TT";
            }

            // Độ rộng cột
            if (c == colTenTT38) resultSheet.Columns[c].ColumnWidth = 35;
            else if (c == colGhiChuDM || c == colGhiChuDG) resultSheet.Columns[c].ColumnWidth = 25;
            else resultSheet.Columns[c].ColumnWidth = 15;

            // Kẻ khung cho cột mới
            if (lastRow >= headerRow + 1)
            {
                Range colRange = resultSheet.Range[resultSheet.Cells[headerRow, c], resultSheet.Cells[lastRow, c]];
                colRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            }
        }

        // Tiêu đề gộp: "Định mức theo Thông tư 38"
        if (colTenTT38 > 0 && colDmDiff > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colTenTT38], resultSheet.Cells[headerRow, colDmDiff]];
            groupHeader.Merge();
            groupHeader.Value2 = "Định mức theo Thông tư 38";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // Tiêu đề gộp: "Đơn giá thẩm định"
        if (colDonGiaChuan > 0 && colChenhLechDG > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colDonGiaChuan], resultSheet.Cells[headerRow, colChenhLechDG]];
            groupHeader.Merge();
            groupHeader.Value2 = "Đơn giá thẩm định";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // Tiêu đề gộp: "Thành tiền thẩm định"
        if (colThanhTienTD > 0 && colChenhLechTT > 0)
        {
            Range groupHeader = resultSheet.Range[resultSheet.Cells[headerRow, colThanhTienTD], resultSheet.Cells[headerRow, colChenhLechTT]];
            groupHeader.Merge();
            groupHeader.Value2 = "Thành tiền thẩm định";
            groupHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            groupHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
            groupHeader.Font.Bold = true;
            groupHeader.Font.Size = 12;
            groupHeader.Font.Name = baseFontName;
            groupHeader.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            groupHeader.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        // ============================================================
        // GHI DỮ LIỆU (Xử lý từ dưới lên trên để chèn dòng không sai lệch)
        // ============================================================
        foreach (var kq in results.OrderByDescending(x => x.DuToan.SoDongExcel))
        {
            var dt = kq.DuToan;
            
            // 1. Tô màu dòng công tác
            Range ctRange = resultSheet.Range[resultSheet.Cells[dt.SoDongExcel, 1], resultSheet.Cells[dt.SoDongExcel, lastCol]];
            
            if (kq.DinhMucChuan == null)
            {
                ctRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 230, 153)); // Cam nhạt
                resultSheet.Cells[dt.SoDongExcel, colGhiChuDM].Value2 = "Mã hiệu không tồn tại trong TT38";
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
            
            // In Tên công tác chuẩn TT38
            if (colTenTT38 > 0)
            {
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Value2 = kq.DinhMucChuan.TenCongTac;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Font.Bold = true;
                resultSheet.Cells[dt.SoDongExcel, colTenTT38].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));
            }

            // 2. Chèn dòng cho các "Hao phí thiếu" (Có trong TT38 nhưng không có trong Dự toán)
            var thieuItems = kq.DanhSachSaiLech.Where(x => x.HaoPhiDuToan == null && x.HaoPhiChuan != null).ToList();
            var loaiHps = new[] { AIE.Core.Enums.LoaiHaoPhi.MAY, AIE.Core.Enums.LoaiHaoPhi.NC, AIE.Core.Enums.LoaiHaoPhi.VL };
            
            foreach (var loai in loaiHps)
            {
                var thieuLoai = thieuItems.Where(x => x.HaoPhiChuan.LoaiHaoPhi == loai).ToList();
                if (thieuLoai.Any())
                {
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

                    for (int i = 0; i < thieuLoai.Count; i++)
                    {
                        int newRow = insertRow + 1 + i;
                        Range insertedRow = (Range)resultSheet.Rows[newRow];
                        insertedRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        
                        Range newRowRange = resultSheet.Range[resultSheet.Cells[newRow, 1], resultSheet.Cells[newRow, lastCol]];
                        newRowRange.Interior.ColorIndex = 0;
                        newRowRange.Font.Bold = false;
                        
                        var sl = thieuLoai[i];
                        if (colTenTT38 > 0) resultSheet.Cells[newRow, colTenTT38].Value2 = sl.HaoPhiChuan.TenHaoPhi;
                        if (colDviTT38 > 0) resultSheet.Cells[newRow, colDviTT38].Value2 = sl.HaoPhiChuan.DonVi;
                        if (colDmTT38 > 0) resultSheet.Cells[newRow, colDmTT38].Value2 = sl.HaoPhiChuan.DinhMuc;
                        
                        // Đơn giá chuẩn cho hao phí thiếu
                        if (colDonGiaChuan > 0 && sl.DonGiaChuan.HasValue)
                            resultSheet.Cells[newRow, colDonGiaChuan].Value2 = sl.DonGiaChuan.Value;
                        
                        // Thành tiền thẩm định cho hao phí thiếu
                        if (colThanhTienTD > 0 && sl.DonGiaChuan.HasValue)
                            resultSheet.Cells[newRow, colThanhTienTD].Value2 = sl.HaoPhiChuan.DinhMuc * sl.DonGiaChuan.Value;
                        
                        resultSheet.Cells[newRow, colGhiChuDM].Value2 = $"[Thiếu hao phí] TT38 có '{sl.HaoPhiChuan.TenHaoPhi}'";
                        resultSheet.Cells[newRow, colGhiChuDM].Font.Color = ColorTranslator.ToOle(Color.Red);

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
                
                if (saiLech != null)
                {
                    if (saiLech.HaoPhiChuan == null)
                    {
                        // Hao phí thừa
                        resultSheet.Cells[rowExcel, colGhiChuDM].Value2 = "Hao phí thừa";
                        resultSheet.Cells[rowExcel, colGhiChuDM].Font.Color = ColorTranslator.ToOle(Color.Red);
                        resultSheet.Cells[rowExcel, colGhiChuDM].Font.Bold = true;
                    }
                    else
                    {
                        // --- Phần Định mức TT38 ---
                        if (colTenTT38 > 0) resultSheet.Cells[rowExcel, colTenTT38].Value2 = saiLech.HaoPhiChuan.TenHaoPhi;
                        if (colDviTT38 > 0) resultSheet.Cells[rowExcel, colDviTT38].Value2 = saiLech.HaoPhiChuan.DonVi;
                        if (colDmTT38 > 0) resultSheet.Cells[rowExcel, colDmTT38].Value2 = saiLech.HaoPhiChuan.DinhMuc;
                        if (colDmDiff > 0)
                        {
                            resultSheet.Cells[rowExcel, colDmDiff].Value2 = saiLech.ChenhLechDinhMuc;
                            if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                                resultSheet.Cells[rowExcel, colDmDiff].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Phần Đơn giá ---
                        if (colDonGiaChuan > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            resultSheet.Cells[rowExcel, colDonGiaChuan].Value2 = saiLech.DonGiaChuan.Value;
                        }
                        if (colChenhLechDG > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal chenhLechDG = hpDuToan.DonGia - saiLech.DonGiaChuan.Value;
                            resultSheet.Cells[rowExcel, colChenhLechDG].Value2 = chenhLechDG;
                            if (Math.Abs(chenhLechDG) > 1)
                                resultSheet.Cells[rowExcel, colChenhLechDG].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Phần Thành tiền ---
                        if (colThanhTienTD > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal ttThamDinh = saiLech.HaoPhiChuan.DinhMuc * saiLech.DonGiaChuan.Value;
                            resultSheet.Cells[rowExcel, colThanhTienTD].Value2 = ttThamDinh;
                        }
                        if (colChenhLechTT > 0 && saiLech.DonGiaChuan.HasValue)
                        {
                            decimal ttThamDinh = saiLech.HaoPhiChuan.DinhMuc * saiLech.DonGiaChuan.Value;
                            decimal ttDuToan = hpDuToan.ThanhTien > 0 ? hpDuToan.ThanhTien : hpDuToan.DinhMuc * hpDuToan.DonGia;
                            decimal chenhLechTT = ttDuToan - ttThamDinh;
                            resultSheet.Cells[rowExcel, colChenhLechTT].Value2 = chenhLechTT;
                            if (Math.Abs(chenhLechTT) > 1)
                                resultSheet.Cells[rowExcel, colChenhLechTT].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }

                        // --- Ghi chú lỗi tổng hợp (định mức + đơn giá) ---
                        // Lỗi từ thẩm định định mức
                        if (!string.IsNullOrEmpty(saiLech.LoaiLoi))
                        {
                            resultSheet.Cells[rowExcel, colGhiChuDM].Value2 = saiLech.LoaiLoi;
                            bool hasSerious = saiLech.LoaiLoi.Contains("Sai") || saiLech.LoaiLoi.Contains("thừa");
                            resultSheet.Cells[rowExcel, colGhiChuDM].Font.Color = hasSerious
                                ? ColorTranslator.ToOle(Color.Red)
                                : ColorTranslator.ToOle(Color.DarkOrange);
                        }

                        // Lỗi đơn giá
                        string loiDonGia = "";
                        if (saiLech.DonGiaChuan.HasValue)
                        {
                            decimal chenhLechDG = hpDuToan.DonGia - saiLech.DonGiaChuan.Value;
                            if (Math.Abs(chenhLechDG) > 1)
                                loiDonGia = "Sai đơn giá";
                        }
                        else if (!string.IsNullOrEmpty(saiLech.HaoPhiChuan?.MaHieuHP))
                        {
                            loiDonGia = "Không có giá chuẩn";
                        }

                        if (!string.IsNullOrEmpty(saiLech.GhiChuDonGia))
                        {
                            loiDonGia = string.IsNullOrEmpty(loiDonGia) ? saiLech.GhiChuDonGia : $"{loiDonGia}. {saiLech.GhiChuDonGia}";
                        }

                        if (!string.IsNullOrEmpty(loiDonGia))
                        {
                            resultSheet.Cells[rowExcel, colGhiChuDG].Value2 = loiDonGia;
                            resultSheet.Cells[rowExcel, colGhiChuDG].Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }
                }
            }
        }

        // ============================================================
        // CỐ ĐỊNH DÒNG TIÊU ĐỀ VÀ ĐỊNH DẠNG CUỐI CÙNG
        // ============================================================
        try
        {
            resultSheet.Activate();
            
            // Cuộn màn hình lên góc trái trên cùng
            resultSheet.Application.ActiveWindow.ScrollRow = 1;
            resultSheet.Application.ActiveWindow.ScrollColumn = 1;
            
            // Cố định dưới dòng headerRow + 1 (tiêu đề có 2 dòng)
            resultSheet.Application.ActiveWindow.SplitRow = headerRow + 1;
            resultSheet.Application.ActiveWindow.SplitColumn = 0;
            resultSheet.Application.ActiveWindow.FreezePanes = true;

            // Wrap Text cho vùng dữ liệu
            if (lastRow > headerRow + 1)
            {
                Range dataRange = resultSheet.Range[resultSheet.Cells[headerRow + 2, 1], resultSheet.Cells[lastRow, colGhiChuDG]];
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
}
