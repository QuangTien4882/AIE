using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Services
{
    public class XuatBangBieuService
    {
        public void Xuat7BangBieu(DuToan duToan)
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null) throw new Exception("Không có Workbook nào đang mở.");

            // Tắt cập nhật màn hình để tăng tốc độ
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;

            try
            {
                Worksheet wsDuToan = null;
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name.StartsWith("DuToan"))
                    {
                        wsDuToan = sheet;
                        break;
                    }
                }
                if (wsDuToan == null) wsDuToan = wb.ActiveSheet as Worksheet;

                // Bảng Hao phí và Cước vận chuyển (Tạo theo thứ tự phụ thuộc công thức)
                XuatBangTongHopNhanCong(wb, duToan);
                XuatBangTongHopCaMay(wb, duToan);
                XuatChietTinhCuocVC(wb, duToan);
                XuatBangTongHopVatLieu(wb, duToan);

                // Cập nhật header DuToan sang format 2 dòng nếu cần
                if (wsDuToan != null)
                {
                    ReformatDuToanHeader(wsDuToan, duToan);
                }

                // Bảng 2: Bảng phân tích đơn giá chi tiết (PhanTich_DonGia) - link các đơn giá vào sheet DuToan
                XuatPhanTichDonGia(wb, duToan, wsDuToan);

                // Bảng 1: Bảng tổng hợp chi phí xây dựng (TH_ChiPhiXD - Bảng 3.8 TT 36)
                XuatBangTongHopChiPhiXayDung(wb, duToan);

                // Bảng 7: Bảng xác định hệ số (HeSo_DieuChinh)
                XuatBangHeSoDieuChinh(wb, duToan);

                // Sắp xếp lại thứ tự sheet theo đúng chuẩn
                SapXepLaiThuTuCacSheet(wb, wsDuToan);

                // Kích hoạt tính toán toàn bộ Workbook để cập nhật các công thức liên kết chéo giữa các sheet
                app.Calculation = XlCalculation.xlCalculationAutomatic;
                try { app.Calculate(); } catch { }

                // Sau khi toàn bộ Workbook đã được tính toán đầy đủ giá trị thực tế, cập nhật lại dòng "Bằng chữ"
                CapNhatDongBangChuSauKhiTinhToan(wb);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Calculation = XlCalculation.xlCalculationAutomatic;
            }
        }

        /// <summary>
        /// Xuất các bảng biểu theo lựa chọn chi tiết của người dùng.
        /// </summary>
        public void XuatCacBangTheoTuyChon(Workbook wb, DuToan duToan, LuaChonXuatExcel opts)
        {
            if (wb == null) throw new Exception("Không có Workbook nào đang mở.");
            if (opts == null || !opts.CoItNhatMotBangDuocChon()) return;

            var app = wb.Application;
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;

            try
            {
                Worksheet wsDuToan = null;
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name.StartsWith("DuToan"))
                    {
                        wsDuToan = sheet;
                        break;
                    }
                }
                if (wsDuToan == null) wsDuToan = wb.ActiveSheet as Worksheet;

                // Xuất theo thứ tự phụ thuộc công thức
                if (opts.XuatTongHopNhanCong && duToan.BangTongHop != null)
                {
                    XuatBangTongHopNhanCong(wb, duToan);
                }
                if (opts.XuatTongHopCaMay && duToan.BangTongHop != null)
                {
                    XuatBangTongHopCaMay(wb, duToan);
                }
                if (opts.XuatChietTinhCuocVC)
                {
                    XuatChietTinhCuocVC(wb, duToan);
                }
                if (opts.XuatTongHopVatLieu && duToan.BangTongHop != null)
                {
                    XuatBangTongHopVatLieu(wb, duToan);
                }
                if (opts.XuatDuToanChiTiet && wsDuToan != null)
                {
                    ReformatDuToanHeader(wsDuToan, duToan);
                }
                if (opts.XuatPhanTichDonGia)
                {
                    XuatPhanTichDonGia(wb, duToan, wsDuToan);
                }
                if (opts.XuatChiPhiXayDung && duToan.ChiPhiXD != null)
                {
                    XuatBangTongHopChiPhiXayDung(wb, duToan);
                }
                if (opts.XuatHeSoDieuChinh)
                {
                    XuatBangHeSoDieuChinh(wb, duToan);
                }
                if (opts.XuatTongHopDuToan && duToan.BangKinhPhi != null)
                {
                    XuatBangTongHopDuToanCongTrinh(wb, duToan);
                }
                if (opts.XuatTongMucDauTu && duToan.BangKinhPhi != null)
                {
                    XuatBangTongMucDauTu(wb, duToan);
                }

                // Sắp xếp lại thứ tự sheet theo đúng chuẩn
                SapXepLaiThuTuCacSheet(wb, wsDuToan);

                // Kích hoạt tính toán toàn bộ Workbook để cập nhật các công thức liên kết chéo giữa các sheet (TH_ChiPhiXD -> TH_DuToan / TongMucDauTu)
                app.Calculation = XlCalculation.xlCalculationAutomatic;
                try { app.Calculate(); } catch { }

                // Sau khi toàn bộ Workbook đã được tính toán đầy đủ giá trị thực tế, cập nhật lại dòng "Bằng chữ"
                CapNhatDongBangChuSauKhiTinhToan(wb);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Calculation = XlCalculation.xlCalculationAutomatic;
                try { app.Calculate(); } catch { }
            }
        }

        public void SapXepLaiThuTuCacSheet(Workbook wb, Worksheet wsDuToan)
        {
            try
            {
                var wsTHChiPhiXD = GetSheetSafe(wb, "TH_ChiPhiXD");
                var wsPhanTich = GetSheetSafe(wb, "PhanTich_DonGia");
                var wsTHVL = GetSheetSafe(wb, "TH_VatLieu");
                var wsTHNC = GetSheetSafe(wb, "TH_NhanCong");
                var wsTHMay = GetSheetSafe(wb, "TH_CaMay");
                var wsCuocVC = GetSheetSafe(wb, "ChietTinh_CuocVC");
                var wsHeSo = GetSheetSafe(wb, "HeSo_DieuChinh");
                var wsTMDT = GetSheetSafe(wb, "TongMucDauTu");
                var wsTHDT = GetSheetSafe(wb, "TH_DuToan");

                // 1. Đặt TH_ChiPhiXD nằm ngay trước sheet DuToan
                if (wsTHChiPhiXD != null)
                {
                    if (wsDuToan != null)
                    {
                        wsTHChiPhiXD.Move(Before: wsDuToan);
                    }
                    else if (wb.Sheets.Count > 0)
                    {
                        wsTHChiPhiXD.Move(Before: wb.Sheets[1]);
                    }
                }

                // 2. Nếu có sheet TongMucDauTu hoặc TH_DuToan thì đặt trước TH_ChiPhiXD
                Worksheet frontTarget = wsTHChiPhiXD ?? wsDuToan;
                if (frontTarget != null)
                {
                    if (wsTHDT != null) wsTHDT.Move(Before: frontTarget);
                    if (wsTMDT != null) wsTMDT.Move(Before: frontTarget);
                }

                // 3. Xếp các sheet sau DuToan theo đúng thứ tự:
                // DuToan -> PhanTich_DonGia -> TH_VatLieu -> TH_NhanCong -> TH_CaMay -> ChietTinh_CuocVC -> HeSo_DieuChinh
                Worksheet prevSheet = wsDuToan ?? wsTHChiPhiXD;
                if (wsPhanTich != null && prevSheet != null)
                {
                    wsPhanTich.Move(After: prevSheet);
                    prevSheet = wsPhanTich;
                }
                if (wsTHVL != null && prevSheet != null)
                {
                    wsTHVL.Move(After: prevSheet);
                    prevSheet = wsTHVL;
                }
                if (wsTHNC != null && prevSheet != null)
                {
                    wsTHNC.Move(After: prevSheet);
                    prevSheet = wsTHNC;
                }
                if (wsTHMay != null && prevSheet != null)
                {
                    wsTHMay.Move(After: prevSheet);
                    prevSheet = wsTHMay;
                }
                if (wsCuocVC != null && prevSheet != null)
                {
                    wsCuocVC.Move(After: prevSheet);
                    prevSheet = wsCuocVC;
                }
                if (wsHeSo != null && prevSheet != null)
                {
                    wsHeSo.Move(After: prevSheet);
                    prevSheet = wsHeSo;
                }
            }
            catch { }
        }

        private static Worksheet GetSheetSafe(Workbook wb, string sheetName)
        {
            if (wb == null || string.IsNullOrEmpty(sheetName)) return null;
            try
            {
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        return sheet;
                }
            }
            catch { }
            return null;
        }

        private static string ChuanHoaChuThuong(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            text = text.Trim();
            // Nếu toàn bộ ký tự chữ cái viết in hoa (ví dụ: "KÊNH ÔNG THU", "ĐÀ NẴNG"), chuyển thành chữ thường kiểu TitleCase ("Kênh Ông Thu", "Đà Nẵng")
            bool hasLetters = text.Any(char.IsLetter);
            bool isAllUpper = hasLetters && text.Where(char.IsLetter).All(char.IsUpper);
            if (isAllUpper)
            {
                var textInfo = new System.Globalization.CultureInfo("vi-VN", false).TextInfo;
                return textInfo.ToTitleCase(text.ToLower());
            }
            return text;
        }

        private static void SetCellSymbolWithSubscript(Worksheet ws, int row, int col, string symbolText)
        {
            if (string.IsNullOrEmpty(symbolText) || ws == null) return;

            string text = symbolText.Trim();
            if (text.Equals("VTM", StringComparison.OrdinalIgnoreCase) || text.Equals("Gtm", StringComparison.OrdinalIgnoreCase))
            {
                text = "GTM";
            }
            else if (text.StartsWith("G_BT", StringComparison.OrdinalIgnoreCase) || text.StartsWith("GBT", StringComparison.OrdinalIgnoreCase))
            {
                text = "GBT, TĐC";
            }
            else if (text.Equals("Gqlda", StringComparison.OrdinalIgnoreCase) || text.Equals("G_QLDA", StringComparison.OrdinalIgnoreCase))
            {
                text = "GQLDA";
            }
            else if (text.Equals("Gtb", StringComparison.OrdinalIgnoreCase) || text.Equals("G_TB", StringComparison.OrdinalIgnoreCase))
            {
                text = "GTB";
            }
            else if (text.Equals("Gtv", StringComparison.OrdinalIgnoreCase) || text.Equals("G_TV", StringComparison.OrdinalIgnoreCase))
            {
                text = "GTV";
            }
            else if (text.Equals("Gk", StringComparison.OrdinalIgnoreCase) || text.Equals("G_K", StringComparison.OrdinalIgnoreCase))
            {
                text = "GK";
            }
            else if (text.Equals("Gdp", StringComparison.OrdinalIgnoreCase) || text.Equals("G_DP", StringComparison.OrdinalIgnoreCase))
            {
                text = "GDP";
            }
            else if (text.Equals("Gxdct", StringComparison.OrdinalIgnoreCase) || text.Equals("G_XDCT", StringComparison.OrdinalIgnoreCase))
            {
                text = "GXDCT";
            }
            else if (text.StartsWith("G_", StringComparison.OrdinalIgnoreCase))
            {
                text = "G" + text.Substring(2);
            }

            Range cell = ws.Cells[row, col];
            cell.Value2 = text;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            if (text.StartsWith("G", StringComparison.OrdinalIgnoreCase) && text.Length > 1)
            {
                try
                {
                    cell.Characters[2, text.Length - 1].Font.Subscript = true;
                }
                catch { }
            }
        }

        private Worksheet CreateOrGetSheet(Workbook wb, string sheetName)
        {
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    sheet.Cells.Clear();
                    return sheet;
                }
            }
            var newSheet = (Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            newSheet.Name = sheetName;
            return newSheet;
        }

        /// <summary>
        /// Kiểm tra và chuyển đổi header sheet DuToan từ format 1 dòng (cũ) sang format 2 dòng (mới).
        /// Format mới: Dòng 4 có "Đơn giá" merged F4:H4, Dòng 5 có "Vật liệu", "Nhân công", "Máy thi công".
        /// </summary>
        public void ReformatDuToanHeader(Worksheet ws, DuToan duToan)
        {
            try
            {
                var app = ws.Application;
                bool oldAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
                try
                {
                    string rawDuAn = !string.IsNullOrEmpty(duToan.TenDuAn) ? duToan.TenDuAn : (!string.IsNullOrEmpty(duToan.TenCongTrinh) ? duToan.TenCongTrinh : "");
                    string rawDiaDiem = !string.IsNullOrEmpty(duToan.DiaDiem) ? duToan.DiaDiem : "";

                    string valA1 = ws.Cells[1, 1].Value2?.ToString() ?? "";
                    string valA2 = ws.Cells[2, 1].Value2?.ToString() ?? "";
                    string valA3 = ws.Cells[3, 1].Value2?.ToString() ?? "";

                    // Đọc lại tên dự án và địa điểm từ các ô nếu duToan chưa có
                    if (string.IsNullOrEmpty(rawDuAn))
                    {
                        if (valA1.Contains(":")) rawDuAn = valA1.Substring(valA1.IndexOf(':') + 1).Trim();
                        else if (valA2.Contains(":")) rawDuAn = valA2.Substring(valA2.IndexOf(':') + 1).Trim();
                    }
                    if (string.IsNullOrEmpty(rawDiaDiem))
                    {
                        if (valA2.Contains("Địa điểm") || valA2.Contains("ĐỊA ĐIỂM")) rawDiaDiem = valA2.Substring(valA2.IndexOf(':') + 1).Trim();
                        else if (valA3.Contains("Địa điểm") || valA3.Contains("ĐỊA ĐIỂM")) rawDiaDiem = valA3.Substring(valA3.IndexOf(':') + 1).Trim();
                    }

                    string tenDuAn = ChuanHoaChuThuong(rawDuAn);
                    if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
                    string diaDiem = ChuanHoaChuThuong(rawDiaDiem);
                    if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";

                    // 1. Kiểm tra xem dòng 1 đã là "BẢNG DỰ TOÁN CHI TIẾT" chưa
                    if (!valA1.ToUpper().Contains("BẢNG DỰ TOÁN"))
                    {
                        // Chèn 1 dòng tại dòng 1 để làm tiêu đề sheet
                        Range row1 = ws.Rows[1];
                        row1.Insert(XlInsertShiftDirection.xlShiftDown);

                        foreach (var hm in duToan.DanhSachHangMuc)
                        {
                            foreach (var ct in hm.DanhSachCongTac)
                            {
                                if (ct.STT >= 1) ct.STT += 1;
                            }
                        }
                    }

                    // 2. Kiểm tra nếu có dòng thừa "Đơn giá" tại dòng 4 (như trường hợp dòng 3 đã có Đơn giá, dòng 4 lại có Đơn giá)
                    string valF3 = ws.Cells[3, 6].Value2?.ToString() ?? "";
                    string valF4 = ws.Cells[4, 6].Value2?.ToString() ?? "";
                    string valF5 = ws.Cells[5, 6].Value2?.ToString() ?? "";
                    if (valF3.ToLower().Contains("đơn giá") && valF4.ToLower().Contains("đơn giá") && valF5.ToLower().Contains("vật liệu"))
                    {
                        // Dòng 4 bị thừa dòng "Đơn giá"! Xóa dòng 4
                        Range row4 = ws.Rows[4];
                        row4.Delete(XlDeleteShiftDirection.xlShiftUp);

                        foreach (var hm in duToan.DanhSachHangMuc)
                        {
                            foreach (var ct in hm.DanhSachCongTac)
                            {
                                if (ct.STT >= 5) ct.STT -= 1;
                            }
                        }
                    }

                    // 3. Kiểm tra xem dòng 5 đã là subheader (Vật liệu) chưa, nếu dòng 5 là dữ liệu (A5 là số) thì chèn dòng 5
                    string checkVL = ws.Cells[5, 6].Value2?.ToString() ?? "";
                    string checkA5 = ws.Cells[5, 1].Value2?.ToString() ?? "";
                    bool a5IsNum = !string.IsNullOrEmpty(checkA5) && int.TryParse(checkA5, out _);
                    if (!checkVL.ToLower().Contains("vật liệu") && !checkVL.ToLower().Equals("vl") && a5IsNum)
                    {
                        Range row5 = ws.Rows[5];
                        row5.Insert(XlInsertShiftDirection.xlShiftDown);
                        foreach (var hm in duToan.DanhSachHangMuc)
                        {
                            foreach (var ct in hm.DanhSachCongTac)
                            {
                                if (ct.STT >= 5) ct.STT += 1;
                            }
                        }
                    }

                    ws.Cells.Font.Name = "Times New Roman";
                    ws.Cells.Font.Size = 12;

                    // Dòng 1: Tiêu đề sheet BẢNG DỰ TOÁN CHI TIẾT
                    var titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 11]];
                    try { titleRange.UnMerge(); } catch { }
                    titleRange.Merge();
                    titleRange.Value2 = "BẢNG DỰ TOÁN CHI TIẾT";
                    titleRange.Font.Bold = true;
                    titleRange.Font.Size = 14;
                    titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    // Dòng 2: Dự án (đồng bộ với TongMucDauTu/TH_DuToan)
                    var duAnRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, 11]];
                    try { duAnRange.UnMerge(); } catch { }
                    duAnRange.Merge();
                    duAnRange.Value2 = $"Dự án: {tenDuAn}";
                    duAnRange.Font.Bold = true;
                    duAnRange.Font.Size = 12;
                    duAnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    duAnRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    // Dòng 3: Địa điểm xây dựng (đồng bộ với TongMucDauTu/TH_DuToan)
                    var diaDiemRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 11]];
                    try { diaDiemRange.UnMerge(); } catch { }
                    diaDiemRange.Merge();
                    diaDiemRange.Value2 = $"Địa điểm xây dựng: {diaDiem}";
                    diaDiemRange.Font.Bold = true;
                    diaDiemRange.Font.Size = 12;
                    diaDiemRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    diaDiemRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    // Dòng 4 & 5: Header 2 dòng chuẩn
                    ws.Cells[4, 1] = "STT";
                    ws.Cells[4, 2] = "Mã hiệu";
                    ws.Cells[4, 3] = "Tên công tác";
                    ws.Cells[4, 4] = "Đơn vị";
                    ws.Cells[4, 5] = "Khối lượng";
                    ws.Cells[4, 6] = "Đơn giá (đồng)";
                    ws.Cells[4, 9] = "Thành tiền (đồng)";

                    ws.Cells[5, 6] = "Vật liệu";
                    ws.Cells[5, 7] = "Nhân công";
                    ws.Cells[5, 8] = "Máy thi công";

                    ws.Cells[5, 9] = "Vật liệu";
                    ws.Cells[5, 10] = "Nhân công";
                    ws.Cells[5, 11] = "Máy thi công";

                    try { ws.Range["A4:A5"].UnMerge(); } catch { }
                    try { ws.Range["B4:B5"].UnMerge(); } catch { }
                    try { ws.Range["C4:C5"].UnMerge(); } catch { }
                    try { ws.Range["D4:D5"].UnMerge(); } catch { }
                    try { ws.Range["E4:E5"].UnMerge(); } catch { }
                    try { ws.Range["F4:H4"].UnMerge(); } catch { }
                    try { ws.Range["I4:K4"].UnMerge(); } catch { }

                    ws.Range["A4:A5"].Merge();
                    ws.Range["B4:B5"].Merge();
                    ws.Range["C4:C5"].Merge();
                    ws.Range["D4:D5"].Merge();
                    ws.Range["E4:E5"].Merge();
                    ws.Range["F4:H4"].Merge();
                    ws.Range["I4:K4"].Merge();

                    Range headerRange = ws.Range[ws.Cells[4, 1], ws.Cells[5, 11]];
                    headerRange.Font.Bold = true;
                    headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 220, 240));
                    headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    // Column widths
                    ((Range)ws.Columns[1]).ColumnWidth = 5;
                    ((Range)ws.Columns[2]).ColumnWidth = 12;
                    ((Range)ws.Columns[3]).ColumnWidth = 42;
                    ((Range)ws.Columns[4]).ColumnWidth = 8;
                    ((Range)ws.Columns[5]).ColumnWidth = 12;
                    ((Range)ws.Columns[6]).ColumnWidth = 14;
                    ((Range)ws.Columns[7]).ColumnWidth = 14;
                    ((Range)ws.Columns[8]).ColumnWidth = 14;
                    ((Range)ws.Columns[9]).ColumnWidth = 16;
                    ((Range)ws.Columns[10]).ColumnWidth = 16;
                    ((Range)ws.Columns[11]).ColumnWidth = 16;

                    // Cập nhật công thức 3 cột Thành tiền cho từng dòng dữ liệu
                    int dataStartRow = 6;
                    int maxDataRow = dataStartRow - 1;
                    for (int r = dataStartRow; r <= 10000; r++)
                    {
                        string mh = ws.Cells[r, 2]?.Value2?.ToString()?.Trim() ?? "";
                        string ten = ws.Cells[r, 3]?.Value2?.ToString()?.Trim() ?? "";
                        if (ten.Equals("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || ten.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                        {
                            break;
                        }
                        if (string.IsNullOrEmpty(mh) && string.IsNullOrEmpty(ten))
                        {
                            break;
                        }
                        maxDataRow = r;
                        ws.Cells[r, 9].Formula = $"=ROUND(E{r}*F{r}, 0)";
                        ws.Cells[r, 10].Formula = $"=ROUND(E{r}*G{r}, 0)";
                        ws.Cells[r, 11].Formula = $"=ROUND(E{r}*H{r}, 0)";
                    }

                    int totalRow = maxDataRow + 1;
                    if (maxDataRow >= dataStartRow)
                    {
                        ws.Cells[totalRow, 3] = "TỔNG CỘNG";
                        ws.Cells[totalRow, 9].Formula = $"=SUM(I6:I{maxDataRow})";
                        ws.Cells[totalRow, 10].Formula = $"=SUM(J6:J{maxDataRow})";
                        ws.Cells[totalRow, 11].Formula = $"=SUM(K6:K{maxDataRow})";
                        ws.Range[ws.Cells[totalRow, 1], ws.Cells[totalRow, 11]].Font.Bold = true;
                        ws.Range[ws.Cells[totalRow, 1], ws.Cells[totalRow, 11]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 245, 252));

                        Range allData = ws.Range[ws.Cells[4, 1], ws.Cells[totalRow, 11]];
                        allData.Borders.LineStyle = XlLineStyle.xlContinuous;
                        ws.Range[$"E6:E{totalRow}"].NumberFormat = "#,##0.000";
                        ws.Range[$"F6:K{totalRow}"].NumberFormat = "#,##0";
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
                    app.DisplayAlerts = oldAlerts;
                }
            }
            catch { /* Bỏ qua lỗi format */ }
        }

        private void SetupHeader(Worksheet ws, string title, int colsCount)
        {
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            var titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, colsCount]];
            titleRange.Merge();
            titleRange.Value2 = title;
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        private void DrawTableBorders(Worksheet ws, int startRow, int startCol, int endRow, int endCol)
        {
            var range = ws.Range[ws.Cells[startRow, startCol], ws.Cells[endRow, endCol]];
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        private void ApplyFreezePanes(Worksheet ws, int splitRow)
        {
            try
            {
                ws.Activate();
                var activeWin = ws.Application.ActiveWindow;
                if (activeWin != null)
                {
                    activeWin.FreezePanes = false;
                    activeWin.SplitRow = splitRow;
                    activeWin.SplitColumn = 0;
                    activeWin.FreezePanes = true;
                }
            }
            catch
            {
                try
                {
                    ws.Activate();
                    ((Range)ws.Cells[splitRow + 1, 1]).Select();
                    ws.Application.ActiveWindow.FreezePanes = true;
                }
                catch { }
            }
        }

        private void CapNhatDongBangChuSauKhiTinhToan(Workbook wb)
        {
            if (wb == null) return;
            try
            {
                string[] targetSheets = new[] { "TongMucDauTu", "TH_DuToan" };
                foreach (var name in targetSheets)
                {
                    Worksheet ws = null;
                    try
                    {
                        foreach (Worksheet s in wb.Sheets)
                        {
                            if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                            {
                                ws = s;
                                break;
                            }
                        }
                    }
                    catch { }
                    if (ws == null) continue;

                    try
                    {
                        int maxRows = Math.Min(100, ws.UsedRange.Rows.Count + ws.UsedRange.Row);
                        for (int r = 1; r <= maxRows; r++)
                        {
                            object c2 = ws.Cells[r, 2]?.Value2;
                            if (c2 == null) continue;
                            string s2 = c2.ToString().Trim().ToUpperInvariant();
                            if (s2.Contains("TỔNG MỨC ĐẦU TƯ XÂY DỰNG") || s2.Contains("TỔNG DỰ TOÁN CÔNG TRÌNH"))
                            {
                                object val = ws.Cells[r, 5]?.Value2;
                                decimal tongSauThue = 0m;
                                if (val is double d) tongSauThue = (decimal)d;
                                else if (val is decimal dec) tongSauThue = dec;
                                else if (val != null)
                                {
                                    if (decimal.TryParse(val.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal pVal) ||
                                        decimal.TryParse(val.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out pVal))
                                    {
                                        tongSauThue = pVal;
                                    }
                                }

                                if (tongSauThue > 0)
                                {
                                    tongSauThue = Math.Round(tongSauThue / 1000m, 0, MidpointRounding.AwayFromZero) * 1000m;
                                    string chuTien = UIHelper.DocSoThanhChu(tongSauThue);
                                    int textRow = r + 1;
                                    var textRange = ws.Range[ws.Cells[textRow, 1], ws.Cells[textRow, 6]];
                                    textRange.Value2 = $"Bằng chữ: {chuTien}.";
                                }
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        public void XuatBangTongHopChiPhiXayDung(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_ChiPhiXD");
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // Dòng 1: Tiêu đề bảng
            var titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 5]];
            titleRange.Merge();
            titleRange.Value2 = "BẢNG TỔNG HỢP CHI PHÍ XÂY DỰNG";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Dòng 2: Dự án (Merge A2:E2, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDuAn = !string.IsNullOrEmpty(duToan.TenDuAn) ? duToan.TenDuAn : (!string.IsNullOrEmpty(duToan.TenCongTrinh) ? duToan.TenCongTrinh : "");
            string tenDuAn = ChuanHoaChuThuong(rawDuAn);
            if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
            var duAnRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, 5]];
            duAnRange.Merge();
            duAnRange.Value2 = $"Dự án: {tenDuAn}";
            duAnRange.Font.Bold = true;
            duAnRange.Font.Size = 12;
            duAnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            duAnRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Dòng 3: Địa điểm xây dựng (Merge A3:E3, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDiaDiem = !string.IsNullOrEmpty(duToan.DiaDiem) ? duToan.DiaDiem : "";
            string diaDiem = ChuanHoaChuThuong(rawDiaDiem);
            if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";
            var diaDiemRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 5]];
            diaDiemRange.Merge();
            diaDiemRange.Value2 = $"Địa điểm xây dựng: {diaDiem}";
            diaDiemRange.Font.Bold = true;
            diaDiemRange.Font.Size = 12;
            diaDiemRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            diaDiemRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Dòng 4: Đơn vị tính góc phải
            ws.Cells[4, 5] = "Đơn vị tính: Đồng";
            ws.Cells[4, 5].Font.Italic = true;
            ws.Cells[4, 5].Font.Size = 12;
            ws.Cells[4, 5].HorizontalAlignment = XlHAlign.xlHAlignRight;

            // Dòng 5: Tiêu đề các cột
            ws.Cells[5, 1] = "TT";
            ws.Cells[5, 2] = "Khoản mục chi phí";
            ws.Cells[5, 3] = "Ký hiệu";
            ws.Cells[5, 4] = "Cách tính";
            ws.Cells[5, 5] = "Giá trị (đồng)";

            var headerRange = ws.Range[ws.Cells[5, 1], ws.Cells[5, 5]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 236, 245));

            // Tìm sheet DuToan để liên kết VL/NC/M
            Worksheet wsDuToan = null;
            foreach (Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase))
                {
                    wsDuToan = sheet;
                    break;
                }
            }
            if (wsDuToan == null) wsDuToan = wb.ActiveSheet as Worksheet;
            string duToanName = wsDuToan != null ? wsDuToan.Name : "DuToan";

            int r = 6;

            // Helper ghi 1 dòng dữ liệu Bảng 3.8
            void AddRow(string tt, string khoiMuc, string kyHieu, string cachTinh, object giaTriOrFormula, bool isBold)
            {
                ws.Cells[r, 1] = tt;
                ws.Cells[r, 2] = khoiMuc;
                SetCellSymbolWithSubscript(ws, r, 3, kyHieu);
                ws.Cells[r, 4] = cachTinh;
                
                if (giaTriOrFormula is string strVal && strVal.StartsWith("=")) {
                    ws.Cells[r, 5].Formula = strVal;
                } else {
                    ws.Cells[r, 5] = giaTriOrFormula;
                }
                
                if (isBold) ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]].Font.Bold = true;
                r++;
            }

            // Kiểm tra đa hạng mục: mỗi hạng mục có ChiPhiXD riêng với tỉ lệ CPC/TT/TL theo LoaiCongTrinh
            var hangMucsCoChiPhi = duToan.DanhSachHangMuc?
                .Where(hm => hm.ChiPhiXD != null && hm.DanhSachCongTac.Count > 0)
                .ToList();
            bool isDaHangMuc = hangMucsCoChiPhi != null && hangMucsCoChiPhi.Count > 1;

            int lastDataRow;

            if (isDaHangMuc)
            {
                // ============ CHẾ ĐỘ ĐA HẠNG MỤC ============
                var inv = System.Globalization.CultureInfo.InvariantCulture;
                var hmGXDTTRows = new List<int>();
                var hmGTGTRows = new List<int>();
                var hmGXDRows = new List<int>();
                var hmLTRows = new List<int>();

                int hmIndex = 0;
                foreach (var hm in hangMucsCoChiPhi)
                {
                    hmIndex++;
                    var cpxd = hm.ChiPhiXD;

                    // Dòng tiêu đề hạng mục
                    var sectionTitle = ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]];
                    sectionTitle.Merge();
                    string tenHM = $"HẠNG MỤC {LapDuToanExcelService.ToRomanNumeral(hmIndex)}: {hm.TenHangMuc}";
                    if (!string.IsNullOrEmpty(hm.LoaiCongTrinh)) tenHM += $" ({hm.LoaiCongTrinh})";
                    sectionTitle.Value2 = tenHM;
                    sectionTitle.Font.Bold = true;
                    sectionTitle.Font.Size = 12;
                    sectionTitle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 220, 250));
                    sectionTitle.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    r++;

                    object linkVL, linkNC, linkM;
                    if (hm.RowIndex > 0)
                    {
                        linkVL = $"='{duToanName}'!I{hm.RowIndex}";
                        linkNC = $"='{duToanName}'!J{hm.RowIndex}";
                        linkM = $"='{duToanName}'!K{hm.RowIndex}";
                    }
                    else
                    {
                        linkVL = (double)hm.TongVL;
                        linkNC = (double)hm.TongNC;
                        linkM = (double)hm.TongMay;
                    }

                    int rowT = r;
                    AddRow("I", "Chi phí trực tiếp", "T", "VL + NC + M", $"=SUM(E{r + 1}:E{r + 3})", true);
                    AddRow("1", "Chi phí vật liệu", "VL", "Σ(KL × ĐG_VL)", linkVL, false);
                    AddRow("2", "Chi phí nhân công", "NC", "Σ(KL × ĐG_NC)", linkNC, false);
                    AddRow("3", "Chi phí máy", "M", "Σ(KL × ĐG_M)", linkM, false);

                    int rowGT = r;
                    AddRow("II", "Chi phí gián tiếp", "GT", "C + TT", $"=E{r + 1}+E{r + 2}", true);
                    AddRow("1", "Chi phí chung", "C", $"T × {cpxd.TiLeCPC}%", $"=ROUND(E{rowT}*{cpxd.TiLeCPC.ToString(inv)}/100, 0)", false);
                    AddRow("2", "Chi phí một số CVKXĐ KL từ TK", "TT", $"T × {cpxd.TiLeTT}%", $"=ROUND(E{rowT}*{cpxd.TiLeTT.ToString(inv)}/100, 0)", false);

                    int rowTL = r;
                    AddRow("III", "Thu nhập chịu thuế tính trước", "TL", $"(T + GT) × {cpxd.TiLeTNCTTT}%", $"=ROUND((E{rowT}+E{rowGT})*{cpxd.TiLeTNCTTT.ToString(inv)}/100, 0)", true);

                    int rowGXDTT = r;
                    hmGXDTTRows.Add(r);
                    AddRow("", "Chi phí xây dựng trước thuế", "GXDTT", "T + GT + TL", $"=E{rowT}+E{rowGT}+E{rowTL}", true);

                    int rowGTGT = r;
                    hmGTGTRows.Add(r);
                    AddRow("IV", "Thuế giá trị gia tăng", "GTGT", $"GXDTT × {cpxd.TiLeGTGT}%", $"=ROUND(E{rowGXDTT}*{cpxd.TiLeGTGT.ToString(inv)}/100, 0)", true);

                    hmGXDRows.Add(r);
                    AddRow("", "CHI PHÍ XÂY DỰNG SAU THUẾ", "GXD", "GXDTT + GTGT", $"=E{rowGXDTT}+E{rowGTGT}", true);

                    hmLTRows.Add(r);
                    AddRow("V", "Chi phí nhà tạm", "LT",
                        $"GXDTT × {cpxd.TiLeNhaTam}% × (1 + {cpxd.TiLeGTGT}%)",
                        $"=ROUND(E{rowGXDTT}*{cpxd.TiLeNhaTam.ToString(inv)}/100*(1+E{rowGTGT}/E{rowGXDTT}), 0)", true);

                    r++; // Dòng trống phân cách
                }

                // TỔNG HỢP TOÀN DỰ ÁN
                var summaryTitle = ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]];
                summaryTitle.Merge();
                summaryTitle.Value2 = "TỔNG HỢP CHI PHÍ XÂY DỰNG TOÀN DỰ ÁN";
                summaryTitle.Font.Bold = true;
                summaryTitle.Font.Size = 13;
                summaryTitle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));
                summaryTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                r++;

                ws.Cells[r, 2] = "Chi phí xây dựng trước thuế";
                SetCellSymbolWithSubscript(ws, r, 3, "GXDTT");
                ws.Cells[r, 4] = "Σ GXDTT các hạng mục";
                ws.Cells[r, 5].Formula = $"={string.Join("+", hmGXDTTRows.Select(x => $"E{x}"))}";
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]].Font.Bold = true;
                r++;

                ws.Cells[r, 2] = "Thuế giá trị gia tăng";
                SetCellSymbolWithSubscript(ws, r, 3, "GTGT");
                ws.Cells[r, 4] = "Σ GTGT các hạng mục";
                ws.Cells[r, 5].Formula = $"={string.Join("+", hmGTGTRows.Select(x => $"E{x}"))}";
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]].Font.Bold = true;
                r++;

                ws.Cells[r, 2] = "CHI PHÍ XÂY DỰNG SAU THUẾ";
                SetCellSymbolWithSubscript(ws, r, 3, "GXD");
                ws.Cells[r, 4] = "GXDTT + GTGT";
                ws.Cells[r, 5].Formula = $"=E{r - 2}+E{r - 1}";
                var gxdTotalRng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]];
                gxdTotalRng.Font.Bold = true;
                gxdTotalRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(230, 244, 234));
                r++;

                ws.Cells[r, 2] = "Chi phí nhà tạm để ở và điều hành thi công";
                SetCellSymbolWithSubscript(ws, r, 3, "LT");
                ws.Cells[r, 4] = "Σ LT các hạng mục";
                ws.Cells[r, 5].Formula = $"={string.Join("+", hmLTRows.Select(x => $"E{x}"))}";
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]].Font.Bold = true;

                lastDataRow = r;
            }
            else
            {
                // ============ CHẾ ĐỘ ĐƠN HẠNG MỤC (giữ nguyên logic cũ) ============
                var kqSingle = duToan.ChiPhiXD;
                if (kqSingle == null) return;

                int duToanTotalRow = 0;
                int totalCongTac = duToan.DanhSachHangMuc?.Sum(hm => hm.DanhSachCongTac.Count) ?? 0;
                if (wsDuToan != null)
                {
                    for (int rSearch = 6; rSearch <= Math.Max(20, 6 + totalCongTac + 5); rSearch++)
                    {
                        string cVal = wsDuToan.Cells[rSearch, 3]?.Value2?.ToString()?.Trim() ?? "";
                        if (cVal.Equals("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || cVal.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                        {
                            duToanTotalRow = rSearch;
                            break;
                        }
                    }
                }
                if (duToanTotalRow == 0) duToanTotalRow = 6 + totalCongTac;

                string linkVL = $"='{duToanName}'!I{duToanTotalRow}";
                string linkNC = $"='{duToanName}'!J{duToanTotalRow}";
                string linkM = $"='{duToanName}'!K{duToanTotalRow}";

                AddRow("I", "Chi phí trực tiếp", "T", "VL + NC + M", "=SUM(E7:E9)", true);
                AddRow("1", "Chi phí vật liệu", "VL", "Σ(KL × ĐG_VL)", linkVL, false);
                AddRow("2", "Chi phí nhân công", "NC", "Σ(KL × ĐG_NC)", linkNC, false);
                AddRow("3", "Chi phí máy", "M", "Σ(KL × ĐG_M)", linkM, false);
                
                AddRow("II", "Chi phí gián tiếp", "GT", "C + TT", "=E11+E12", true);
                AddRow("1", "Chi phí chung", "C", $"T × {kqSingle.TiLeCPC}%", $"=ROUND(E6*{kqSingle.TiLeCPC.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100, 0)", false);
                AddRow("2", "Chi phí một số công việc không xác định được KL từ TK", "TT", $"T × {kqSingle.TiLeTT}%", $"=ROUND(E6*{kqSingle.TiLeTT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100, 0)", false);
                
                AddRow("III", "Thu nhập chịu thuế tính trước", "TL", $"(T + GT) × {kqSingle.TiLeTNCTTT}%", $"=ROUND((E6+E10)*{kqSingle.TiLeTNCTTT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100, 0)", true);
                AddRow("", "Chi phí xây dựng trước thuế", "GXDTT", "T + GT + TL", "=E6+E10+E13", true);
                AddRow("IV", "Thuế giá trị gia tăng", "GTGT", $"GXDTT × {kqSingle.TiLeGTGT}%", $"=ROUND(E14*{kqSingle.TiLeGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100, 0)", true);
                AddRow("", "CHI PHÍ XÂY DỰNG SAU THUẾ", "GXD", "GXDTT + GTGT", "=E14+E15", true);
                AddRow("V", "Chi phí nhà tạm để ở và điều hành thi công", "LT", $"GXDTT × {kqSingle.TiLeNhaTam}% × (1 + {kqSingle.TiLeGTGT}%)", $"=ROUND(E14*{kqSingle.TiLeNhaTam.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100*(1+E15/E14), 0)", true);

                lastDataRow = 17;
            }

            // Vẽ viền bảng tổng hợp chi phí xây dựng
            DrawTableBorders(ws, 5, 1, lastDataRow, 5);
            ws.Range["E:E"].NumberFormat = "#,##0";
            ws.Columns.AutoFit();
            ApplyFreezePanes(ws, 5);

            // Thêm chữ ký Người lập và Người chủ trì theo Bảng 3.8
            int signRow = lastDataRow + 2;
            ws.Cells[signRow, 2] = "NGƯỜI LẬP";
            ws.Cells[signRow, 2].Font.Bold = true;
            ws.Cells[signRow, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Cells[signRow, 5] = "NGƯỜI CHỦ TRÌ";
            ws.Cells[signRow, 5].Font.Bold = true;
            ws.Cells[signRow, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Cells[signRow + 1, 2] = "(Ký, họ tên)";
            ws.Cells[signRow + 1, 2].Font.Italic = true;
            ws.Cells[signRow + 1, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Cells[signRow + 1, 5] = "(Ký, họ tên)";
            ws.Cells[signRow + 1, 5].Font.Italic = true;
            ws.Cells[signRow + 1, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Di chuyển sheet TH_ChiPhiXD nằm ngay phía trước sheet DuToan
            try
            {
                if (wsDuToan != null) ws.Move(Before: wsDuToan);
                else if (wb.Sheets.Count > 1) ws.Move(Before: wb.Sheets[1]);
            }
            catch { }
        }

        public void XuatBangTongHopDuToan(Workbook wb, DuToan duToan)
        {
            XuatBangTongHopChiPhiXayDung(wb, duToan);
        }

        public void XuatBangTongHopDuToanCongTrinh(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_DuToan");
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // Dòng 1: Tiêu đề bảng
            var titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 6]];
            titleRange.Merge();
            titleRange.Value2 = "BẢNG TỔNG HỢP DỰ TOÁN CÔNG TRÌNH";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 2: Dự án (Merge A2:F2, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDuAn = !string.IsNullOrEmpty(duToan.TenDuAn) ? duToan.TenDuAn : (!string.IsNullOrEmpty(duToan.TenCongTrinh) ? duToan.TenCongTrinh : "");
            string tenDuAn = ChuanHoaChuThuong(rawDuAn);
            if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
            var duAnRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, 6]];
            duAnRange.Merge();
            duAnRange.Value2 = $"Dự án: {tenDuAn}";
            duAnRange.Font.Bold = true;
            duAnRange.Font.Size = 12;
            duAnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 3: Địa điểm xây dựng (Merge A3:F3, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDiaDiem = !string.IsNullOrEmpty(duToan.DiaDiem) ? duToan.DiaDiem : "";
            string diaDiem = ChuanHoaChuThuong(rawDiaDiem);
            if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";
            var diaDiemRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 6]];
            diaDiemRange.Merge();
            diaDiemRange.Value2 = $"Địa điểm xây dựng: {diaDiem}";
            diaDiemRange.Font.Bold = true;
            diaDiemRange.Font.Size = 12;
            diaDiemRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 4: Đơn vị tính góc phải
            ws.Cells[4, 6] = "Đơn vị tính: Đồng";
            ws.Cells[4, 6].Font.Italic = true;
            ws.Cells[4, 6].Font.Size = 12;
            ws.Cells[4, 6].HorizontalAlignment = XlHAlign.xlHAlignRight;

            // Dòng 5: Tiêu đề các cột
            ws.Cells[5, 1] = "STT";
            ws.Cells[5, 2] = "NỘI DUNG CHI PHÍ";
            ws.Cells[5, 3] = "GIÁ TRỊ TRƯỚC THUẾ";
            ws.Cells[5, 4] = "THUẾ GTGT";
            ws.Cells[5, 5] = "GIÁ TRỊ SAU THUẾ";
            ws.Cells[5, 6] = "KÝ HIỆU";

            // Dòng 6: Đánh số cột [1] [2] [3] [4] [5] [6]
            ws.Cells[6, 1] = "[1]";
            ws.Cells[6, 2] = "[2]";
            ws.Cells[6, 3] = "[3]";
            ws.Cells[6, 4] = "[4]";
            ws.Cells[6, 5] = "[5]";
            ws.Cells[6, 6] = "[6]";

            var headerRange = ws.Range[ws.Cells[5, 1], ws.Cells[6, 6]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));

            ws.Range["A:A"].NumberFormat = "@";

            bool hasChiPhiXD = false;
            foreach (Worksheet sh in wb.Sheets)
            {
                if (sh.Name == "TH_ChiPhiXD") { hasChiPhiXD = true; break; }
            }

            var model = duToan.BangKinhPhi ?? DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: !string.IsNullOrEmpty(duToan.LoaiCongTrinh) ? duToan.LoaiCongTrinh : "Dân dụng",
                capCT: !string.IsNullOrEmpty(duToan.CapCongTrinh) ? duToan.CapCongTrinh : "Cấp III",
                soBuocTK: 2,
                chiPhiXD: duToan.ChiPhiXD?.G ?? 10_000_000_000m,
                chiPhiTB: duToan.ChiPhiThietBi
            );

            int r = 7;

            // 1. Chi phí xây dựng (Tổng nhóm 1)
            int rowXD_Tong = r;
            ws.Cells[r, 1] = "'1";
            ws.Cells[r, 2] = "Chi phí xây dựng";
            SetCellSymbolWithSubscript(ws, r, 6, "GXD");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            // 1.1. Chi phí xây dựng (G_XD - Không bao gồm nhà tạm)
            int rowXD_Con = r;
            ws.Cells[r, 1] = "'1.1";
            ws.Cells[r, 2] = "- Chi phí xây dựng";
            if (hasChiPhiXD)
            {
                ws.Cells[r, 3].Formula = "='TH_ChiPhiXD'!E14";
            }
            else
            {
                ws.Cells[r, 3] = (double)model.ChiPhiXDTruocThue;
            }
            var itemXD = model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
            decimal vatXD = itemXD != null ? itemXD.ThueSuatGTGT : (duToan.ChiPhiXD != null && duToan.ChiPhiXD.TiLeGTGT > 0 ? duToan.ChiPhiXD.TiLeGTGT / 100m : 0.10m);
            string vatXDStr = vatXD.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatXDStr}, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gxd");
            r++;

            // 1.2. Chi phí nhà tạm để ở và điều hành thi công
            int rowNT_Con = r;
            ws.Cells[r, 1] = "'1.2";
            ws.Cells[r, 2] = "- Chi phí nhà tạm để ở và điều hành thi công";
            var itemNT = model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
            decimal vatNT = itemNT != null ? itemNT.ThueSuatGTGT : vatXD;
            string vatNTStr = vatNT.ToString(System.Globalization.CultureInfo.InvariantCulture);

            if (hasChiPhiXD)
            {
                if (duToan.ChiPhiXD != null && duToan.ChiPhiXD.TiLeNhaTam > 0)
                {
                    string tlNT = (duToan.ChiPhiXD.TiLeNhaTam / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    ws.Cells[r, 3].Formula = $"=ROUND('TH_ChiPhiXD'!E14 * {tlNT}, 0)";
                }
                else
                {
                    ws.Cells[r, 3].Formula = $"=ROUND('TH_ChiPhiXD'!E17 / (1 + {vatNTStr}), 0)";
                }
            }
            else
            {
                ws.Cells[r, 3] = (double)model.ChiPhiNhaTamTruocThue;
            }
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatNTStr}, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gnt");
            r++;

            // Dòng tổng nhóm 1
            ws.Cells[rowXD_Tong, 3].Formula = $"=C{rowXD_Con}+C{rowNT_Con}";
            ws.Cells[rowXD_Tong, 4].Formula = $"=D{rowXD_Con}+D{rowNT_Con}";
            ws.Cells[rowXD_Tong, 5].Formula = $"=ROUND(E{rowXD_Con}+E{rowNT_Con}, -3)";

            bool hasTB = model.ChiPhiTBTruocThue > 0;
            int rowTB = 0;

            if (hasTB)
            {
                // 2. Chi phí thiết bị
                rowTB = r;
                ws.Cells[r, 1] = "'2";
                ws.Cells[r, 2] = "Chi phí thiết bị";
                ws.Cells[r, 3] = (double)model.ChiPhiTBTruocThue;
                var itemTB = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                decimal vatTB = itemTB != null ? itemTB.ThueSuatGTGT : 0.10m;
                string vatTBStr = vatTB.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatTBStr}, 0)";
                ws.Cells[r, 5].Formula = $"=ROUND(C{r}+D{r}, -3)";
                SetCellSymbolWithSubscript(ws, r, 6, "GTB");
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
                r++;
            }

            // Chi phí quản lý dự án (Cơ sở tính: Chi phí xây dựng CHƯA có nhà tạm + Thiết bị nếu có)
            int rowQLDA = r;
            var itemQLDA = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.QuanLyDuAn) ?? new ChiPhiKinhPhiItem { TyLePhanTram = 3.25m, HeSoDieuChinh = 1.0m };
            string qldaTyLe = (itemQLDA.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
            string qldaHeSo = itemQLDA.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 1] = hasTB ? "'3" : "'2";
            ws.Cells[r, 2] = "Chi phí quản lý dự án";
            string qldaBase = hasTB ? $"(C{rowXD_Con} + C{rowTB})" : $"C{rowXD_Con}";
            ws.Cells[r, 3].Formula = $"=ROUND({qldaTyLe} * {qldaBase} * {qldaHeSo}, 0)";
            decimal vatQLDA = itemQLDA.ThueSuatGTGT;
            string vatQLDAStr = vatQLDA.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatQLDAStr}, 0)";
            ws.Cells[r, 5].Formula = $"=ROUND(C{r}+D{r}, -3)";
            SetCellSymbolWithSubscript(ws, r, 6, "GQLDA");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            // Chi phí tư vấn đầu tư xây dựng (Cơ sở tính: Chi phí xây dựng con C{rowXD_Con})
            int rowTVGroup = r;
            ws.Cells[r, 1] = hasTB ? "'4" : "'3";
            ws.Cells[r, 2] = "Chi phí tư vấn đầu tư xây dựng";
            SetCellSymbolWithSubscript(ws, r, 6, "GTV");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            int startTV = r;
            var tvItems = model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.TuVanDauTuXD && (hasTB || x.CoSoTinh != CoSoTinhChiPhi.ChiPhiThietBi)).ToList();
            int tvIdx = 1;
            string tvPrefix = hasTB ? "'4." : "'3.";
            foreach (var item in tvItems)
            {
                ws.Cells[r, 1] = $"{tvPrefix}{tvIdx++}";
                ws.Cells[r, 2] = "- " + item.TenChiPhi;
                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    string baseCell = item.CoSoTinh == CoSoTinhChiPhi.ChiPhiXayDung ? $"C{rowXD_Con}" : (item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi ? (hasTB ? $"C{rowTB}" : $"C{rowXD_Con}") : (hasTB ? $"(C{rowXD_Con}+C{rowTB})" : $"C{rowXD_Con}"));
                    string tlStr = (item.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string hsStr = item.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    ws.Cells[r, 3].Formula = $"=ROUND({tlStr} * {baseCell} * {hsStr}, 0)";
                }
                else
                {
                    ws.Cells[r, 3] = (double)item.GiaTriTruocThue;
                }
                string vatStr = item.ThueSuatGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatStr}, 0)";
                ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
                SetCellSymbolWithSubscript(ws, r, 6, item.KyHieu);
                r++;
            }
            int endTV = r - 1;
            if (endTV >= startTV)
            {
                ws.Cells[rowTVGroup, 3].Formula = $"=SUM(C{startTV}:C{endTV})";
                ws.Cells[rowTVGroup, 4].Formula = $"=SUM(D{startTV}:D{endTV})";
                ws.Cells[rowTVGroup, 5].Formula = $"=ROUND(SUM(E{startTV}:E{endTV}), -3)";
            }
            else
            {
                ws.Cells[rowTVGroup, 3] = 0; ws.Cells[rowTVGroup, 4] = 0; ws.Cells[rowTVGroup, 5] = 0;
            }

            // Chi phí khác (Cơ sở tính: Chi phí xây dựng con C{rowXD_Con})
            int rowKGroup = r;
            ws.Cells[r, 1] = hasTB ? "'5" : "'4";
            ws.Cells[r, 2] = "Chi phí khác";
            SetCellSymbolWithSubscript(ws, r, 6, "GK");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            int startK = r;
            var kItems = model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiKhac && (hasTB || x.CoSoTinh != CoSoTinhChiPhi.ChiPhiThietBi)).ToList();
            int kIdx = 1;
            string kPrefix = hasTB ? "'5." : "'4.";
            foreach (var item in kItems)
            {
                ws.Cells[r, 1] = $"{kPrefix}{kIdx++}";
                ws.Cells[r, 2] = "- " + item.TenChiPhi;
                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    string baseCell = item.CoSoTinh == CoSoTinhChiPhi.ChiPhiXayDung ? $"C{rowXD_Con}" : (item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi ? (hasTB ? $"C{rowTB}" : $"C{rowXD_Con}") : (hasTB ? $"(C{rowXD_Con}+C{rowTB})" : $"C{rowXD_Con}"));
                    string tlStr = (item.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string hsStr = item.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (item.MinValue.HasValue && item.MaxValue.HasValue)
                    {
                        string minStr = item.MinValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        string maxStr = item.MaxValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        ws.Cells[r, 3].Formula = $"=ROUND(MAX({minStr}, MIN({maxStr}, {tlStr} * {baseCell} * {hsStr})), 0)";
                    }
                    else if (item.MinValue.HasValue)
                    {
                        string minStr = item.MinValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        ws.Cells[r, 3].Formula = $"=ROUND(MAX({minStr}, {tlStr} * {baseCell} * {hsStr}), 0)";
                    }
                    else
                    {
                        ws.Cells[r, 3].Formula = $"=ROUND({tlStr} * {baseCell} * {hsStr}, 0)";
                    }
                }
                else
                {
                    ws.Cells[r, 3] = (double)item.GiaTriTruocThue;
                }
                string vatStr = item.ThueSuatGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatStr}, 0)";
                ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
                SetCellSymbolWithSubscript(ws, r, 6, item.KyHieu);
                r++;
            }
            int endK = r - 1;
            if (endK >= startK)
            {
                ws.Cells[rowKGroup, 3].Formula = $"=SUM(C{startK}:C{endK})";
                ws.Cells[rowKGroup, 4].Formula = $"=SUM(D{startK}:D{endK})";
                ws.Cells[rowKGroup, 5].Formula = $"=ROUND(SUM(E{startK}:E{endK}), -3)";
            }
            else
            {
                ws.Cells[rowKGroup, 3] = 0; ws.Cells[rowKGroup, 4] = 0; ws.Cells[rowKGroup, 5] = 0;
            }

            // Chi phí dự phòng (G_DP)
            int rowDPGroup = r;
            ws.Cells[r, 1] = hasTB ? "'6" : "'5";
            ws.Cells[r, 2] = "Chi phí dự phòng";
            SetCellSymbolWithSubscript(ws, r, 6, "GDP");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            var itemDP = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong) ?? new ChiPhiKinhPhiItem { TyLePhanTram = 5.0m };
            string dpTyLe = (itemDP.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
            int rowDP1 = r;
            ws.Cells[r, 1] = hasTB ? "'6.1" : "'5.1";
            ws.Cells[r, 2] = "- Chi phí dự phòng yếu tố khối lượng phát sinh (Gdp1)";
            string dpBase = hasTB
                ? $"(C{rowXD_Tong} + C{rowTB} + C{rowQLDA} + C{rowTVGroup} + C{rowKGroup})"
                : $"(C{rowXD_Tong} + C{rowQLDA} + C{rowTVGroup} + C{rowKGroup})";
            ws.Cells[r, 3].Formula = $"=ROUND({dpTyLe} * {dpBase}, 0)";
            ws.Cells[r, 4].Formula = $"=ROUND(C{r}*0.1, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gdp1");
            r++;

            int rowDP2 = r;
            ws.Cells[r, 1] = hasTB ? "'6.2" : "'5.2";
            ws.Cells[r, 2] = "- Chi phí dự phòng yếu tố trượt giá (Gdp2)";
            ws.Cells[r, 3] = 0;
            ws.Cells[r, 4] = 0;
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gdp2");
            r++;

            ws.Cells[rowDPGroup, 3].Formula = $"=C{rowDP1}+C{rowDP2}";
            ws.Cells[rowDPGroup, 4].Formula = $"=D{rowDP1}+D{rowDP2}";
            ws.Cells[rowDPGroup, 5].Formula = $"=ROUND(E{rowDP1}+E{rowDP2}, -3)";

            // Dòng Tổng cộng Dự toán công trình
            int grandRow = r;
            ws.Cells[r, 2] = "TỔNG CỘNG DỰ TOÁN CÔNG TRÌNH";
            string grandSumC = hasTB
                ? $"=C{rowXD_Tong}+C{rowTB}+C{rowQLDA}+C{rowTVGroup}+C{rowKGroup}+C{rowDPGroup}"
                : $"=C{rowXD_Tong}+C{rowQLDA}+C{rowTVGroup}+C{rowKGroup}+C{rowDPGroup}";
            string grandSumD = hasTB
                ? $"=D{rowXD_Tong}+D{rowTB}+D{rowQLDA}+D{rowTVGroup}+D{rowKGroup}+D{rowDPGroup}"
                : $"=D{rowXD_Tong}+D{rowQLDA}+D{rowTVGroup}+D{rowKGroup}+D{rowDPGroup}";
            string grandSumE = hasTB
                ? $"=ROUND(E{rowXD_Tong}+E{rowTB}+E{rowQLDA}+E{rowTVGroup}+E{rowKGroup}+E{rowDPGroup}, -3)"
                : $"=ROUND(E{rowXD_Tong}+E{rowQLDA}+E{rowTVGroup}+E{rowKGroup}+E{rowDPGroup}, -3)";
            ws.Cells[r, 3].Formula = grandSumC;
            ws.Cells[r, 4].Formula = grandSumD;
            ws.Cells[r, 5].Formula = grandSumE;
            SetCellSymbolWithSubscript(ws, r, 6, "GXDCT");

            var grandRng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]];
            grandRng.Font.Bold = true;
            grandRng.Font.Size = 12;
            grandRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));

            // Dòng Bằng chữ (Merge A..F, căn giữa, in nghiêng)
            int textRow = grandRow + 1;
            try { ws.Calculate(); } catch { }
            object valGrand = ws.Cells[grandRow, 5]?.Value2;
            decimal tongSauThue = 0m;
            if (valGrand is double d) tongSauThue = (decimal)d;
            else if (valGrand is decimal dec) tongSauThue = dec;
            else if (valGrand != null)
            {
                if (decimal.TryParse(valGrand.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal pVal) ||
                    decimal.TryParse(valGrand.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out pVal))
                {
                    tongSauThue = pVal;
                }
            }
            if (tongSauThue <= 0)
            {
                tongSauThue = model.TongSauThue;
            }
            tongSauThue = Math.Round(tongSauThue / 1000m, 0, MidpointRounding.AwayFromZero) * 1000m;
            string chuTien = UIHelper.DocSoThanhChu(tongSauThue);
            var textRange = ws.Range[ws.Cells[textRow, 1], ws.Cells[textRow, 6]];
            textRange.Merge();
            textRange.Value2 = $"Bằng chữ: {chuTien}.";
            textRange.Font.Italic = true;
            textRange.Font.Bold = false;
            textRange.Font.Size = 12;
            textRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            textRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            DrawTableBorders(ws, 5, 1, textRow, 6);
            ws.Range[$"C7:E{grandRow}"].NumberFormat = "#,##0;-#,##0;\"-\"";
            ws.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[6].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();

            // Cố định dòng 6 (Freeze Panes) để luôn nhìn thấy tiêu đề khi cuộn dọc
            ApplyFreezePanes(ws, 6);

            // Yêu cầu 3: Di chuyển sheet TH_DuToan nằm ngay phía trước sheet TH_ChiPhiXD
            try
            {
                Worksheet wsTarget = GetSheetSafe(wb, "TH_ChiPhiXD");
                if (wsTarget == null)
                {
                    foreach (Worksheet sh in wb.Sheets)
                    {
                        if (sh.Name.StartsWith("DuToan")) { wsTarget = sh; break; }
                    }
                }
                if (wsTarget != null)
                {
                    ws.Move(Before: wsTarget);
                }
                else if (wb.Sheets.Count > 1)
                {
                    ws.Move(Before: wb.Sheets[1]);
                }
            }
            catch { }
        }

        public void XuatBangTongMucDauTu(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TongMucDauTu");
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // Dòng 1: Tiêu đề bảng
            var titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 6]];
            titleRange.Merge();
            titleRange.Value2 = "BẢNG TỔNG HỢP TỔNG MỨC ĐẦU TƯ XÂY DỰNG";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;
            titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 2: Dự án (Merge A2:F2, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDuAn = !string.IsNullOrEmpty(duToan.TenDuAn) ? duToan.TenDuAn : (!string.IsNullOrEmpty(duToan.TenCongTrinh) ? duToan.TenCongTrinh : "");
            string tenDuAn = ChuanHoaChuThuong(rawDuAn);
            if (string.IsNullOrEmpty(tenDuAn)) tenDuAn = "................................................................";
            var duAnRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, 6]];
            duAnRange.Merge();
            duAnRange.Value2 = $"Dự án: {tenDuAn}";
            duAnRange.Font.Bold = true;
            duAnRange.Font.Size = 12;
            duAnRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 3: Địa điểm xây dựng (Merge A3:F3, căn giữa, in đậm, chữ thường chuẩn TitleCase)
            string rawDiaDiem = !string.IsNullOrEmpty(duToan.DiaDiem) ? duToan.DiaDiem : "";
            string diaDiem = ChuanHoaChuThuong(rawDiaDiem);
            if (string.IsNullOrEmpty(diaDiem)) diaDiem = "................................................................";
            var diaDiemRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 6]];
            diaDiemRange.Merge();
            diaDiemRange.Value2 = $"Địa điểm xây dựng: {diaDiem}";
            diaDiemRange.Font.Bold = true;
            diaDiemRange.Font.Size = 12;
            diaDiemRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // Dòng 4: Đơn vị tính bên phải
            ws.Cells[4, 6] = "Đơn vị tính: Đồng";
            ws.Cells[4, 6].Font.Italic = true;
            ws.Cells[4, 6].Font.Size = 12;
            ws.Cells[4, 6].HorizontalAlignment = XlHAlign.xlHAlignRight;

            // Dòng 5: Tiêu đề các cột
            ws.Cells[5, 1] = "STT";
            ws.Cells[5, 2] = "NỘI DUNG CHI PHÍ";
            ws.Cells[5, 3] = "GIÁ TRỊ TRƯỚC THUẾ";
            ws.Cells[5, 4] = "THUẾ GTGT";
            ws.Cells[5, 5] = "GIÁ TRỊ SAU THUẾ";
            ws.Cells[5, 6] = "KÝ HIỆU";

            // Dòng 6: Đánh số cột [1] [2] [3] [4] [5] [6]
            ws.Cells[6, 1] = "[1]";
            ws.Cells[6, 2] = "[2]";
            ws.Cells[6, 3] = "[3]";
            ws.Cells[6, 4] = "[4]";
            ws.Cells[6, 5] = "[5]";
            ws.Cells[6, 6] = "[6]";

            var headerRange = ws.Range[ws.Cells[5, 1], ws.Cells[6, 6]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));

            ws.Range["A:A"].NumberFormat = "@";

            bool hasChiPhiXD = false;
            foreach (Worksheet sh in wb.Sheets)
            {
                if (sh.Name == "TH_ChiPhiXD") { hasChiPhiXD = true; break; }
            }

            var model = duToan.BangKinhPhi ?? DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                loaiCT: !string.IsNullOrEmpty(duToan.LoaiCongTrinh) ? duToan.LoaiCongTrinh : "Dân dụng",
                capCT: !string.IsNullOrEmpty(duToan.CapCongTrinh) ? duToan.CapCongTrinh : "Cấp III",
                soBuocTK: 2,
                chiPhiXD: duToan.ChiPhiXD?.G ?? 10_000_000_000m,
                chiPhiTB: duToan.ChiPhiThietBi
            );

            int r = 7;

            // 1. Chi phí bồi thường, hỗ trợ và tái định cư
            int rowBT = r;
            ws.Cells[r, 1] = "'1";
            ws.Cells[r, 2] = "Chi phí bồi thường, hỗ trợ và tái định cư";
            ws.Cells[r, 3] = (double)model.ChiPhiBTTruocThue;
            ws.Cells[r, 4] = 0;
            ws.Cells[r, 5].Formula = $"=ROUND(C{r}+D{r}, -3)";
            SetCellSymbolWithSubscript(ws, r, 6, "GBT, TĐC");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            // 2. Chi phí xây dựng (Tổng nhóm 2)
            int rowXD_Tong = r;
            ws.Cells[r, 1] = "'2";
            ws.Cells[r, 2] = "Chi phí xây dựng";
            SetCellSymbolWithSubscript(ws, r, 6, "GXD");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            // 2.1. Chi phí xây dựng (G_XD - Không bao gồm nhà tạm)
            int rowXD_Con = r;
            ws.Cells[r, 1] = "'2.1";
            ws.Cells[r, 2] = "- Chi phí xây dựng";
            if (hasChiPhiXD)
            {
                ws.Cells[r, 3].Formula = "='TH_ChiPhiXD'!E14";
            }
            else
            {
                ws.Cells[r, 3] = (double)model.ChiPhiXDTruocThue;
            }
            var itemXD = model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
            decimal vatXD = itemXD != null ? itemXD.ThueSuatGTGT : (duToan.ChiPhiXD != null && duToan.ChiPhiXD.TiLeGTGT > 0 ? duToan.ChiPhiXD.TiLeGTGT / 100m : 0.10m);
            string vatXDStr = vatXD.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatXDStr}, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gxd");
            r++;

            // 2.2. Chi phí nhà tạm để ở và điều hành thi công
            int rowNT_Con = r;
            ws.Cells[r, 1] = "'2.2";
            ws.Cells[r, 2] = "- Chi phí nhà tạm để ở và điều hành thi công";
            var itemNT = model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
            decimal vatNT = itemNT != null ? itemNT.ThueSuatGTGT : vatXD;
            string vatNTStr = vatNT.ToString(System.Globalization.CultureInfo.InvariantCulture);

            if (hasChiPhiXD)
            {
                if (duToan.ChiPhiXD != null && duToan.ChiPhiXD.TiLeNhaTam > 0)
                {
                    string tlNT = (duToan.ChiPhiXD.TiLeNhaTam / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    ws.Cells[r, 3].Formula = $"=ROUND('TH_ChiPhiXD'!E14 * {tlNT}, 0)";
                }
                else
                {
                    ws.Cells[r, 3].Formula = $"=ROUND('TH_ChiPhiXD'!E17 / (1 + {vatNTStr}), 0)";
                }
            }
            else
            {
                ws.Cells[r, 3] = (double)model.ChiPhiNhaTamTruocThue;
            }
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatNTStr}, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gnt");
            r++;

            // Dòng tổng nhóm 2
            ws.Cells[rowXD_Tong, 3].Formula = $"=C{rowXD_Con}+C{rowNT_Con}";
            ws.Cells[rowXD_Tong, 4].Formula = $"=D{rowXD_Con}+D{rowNT_Con}";
            ws.Cells[rowXD_Tong, 5].Formula = $"=ROUND(E{rowXD_Con}+E{rowNT_Con}, -3)";

            bool hasTB = model.ChiPhiTBTruocThue > 0;
            int rowTB = 0;

            if (hasTB)
            {
                // 3. Chi phí thiết bị
                rowTB = r;
                ws.Cells[r, 1] = "'3";
                ws.Cells[r, 2] = "Chi phí thiết bị";
                ws.Cells[r, 3] = (double)model.ChiPhiTBTruocThue;
                var itemTB = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                decimal vatTB = itemTB != null ? itemTB.ThueSuatGTGT : 0.10m;
                string vatTBStr = vatTB.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatTBStr}, 0)";
                ws.Cells[r, 5].Formula = $"=ROUND(C{r}+D{r}, -3)";
                SetCellSymbolWithSubscript(ws, r, 6, "GTB");
                ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
                r++;
            }

            // Chi phí quản lý dự án (Cơ sở tính: Chi phí xây dựng CHƯA có nhà tạm + Thiết bị nếu có)
            int rowQLDA = r;
            var itemQLDA = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.QuanLyDuAn) ?? new ChiPhiKinhPhiItem { TyLePhanTram = 3.25m, HeSoDieuChinh = 1.0m };
            string qldaTyLe = (itemQLDA.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
            string qldaHeSo = itemQLDA.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 1] = hasTB ? "'4" : "'3";
            ws.Cells[r, 2] = "Chi phí quản lý dự án";
            string qldaBase = hasTB ? $"(C{rowXD_Con} + C{rowTB})" : $"C{rowXD_Con}";
            ws.Cells[r, 3].Formula = $"=ROUND({qldaTyLe} * {qldaBase} * {qldaHeSo}, 0)";
            decimal vatQLDA = itemQLDA.ThueSuatGTGT;
            string vatQLDAStr = vatQLDA.ToString(System.Globalization.CultureInfo.InvariantCulture);
            ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatQLDAStr}, 0)";
            ws.Cells[r, 5].Formula = $"=ROUND(C{r}+D{r}, -3)";
            SetCellSymbolWithSubscript(ws, r, 6, "GQLDA");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            // Chi phí tư vấn đầu tư xây dựng (Cơ sở tính: Chi phí xây dựng con C{rowXD_Con})
            int rowTVGroup = r;
            ws.Cells[r, 1] = hasTB ? "'5" : "'4";
            ws.Cells[r, 2] = "Chi phí tư vấn đầu tư xây dựng";
            SetCellSymbolWithSubscript(ws, r, 6, "GTV");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            int startTV = r;
            var tvItems = model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.TuVanDauTuXD && (hasTB || x.CoSoTinh != CoSoTinhChiPhi.ChiPhiThietBi)).ToList();
            int tvIdx = 1;
            string tvPrefix = hasTB ? "'5." : "'4.";
            foreach (var item in tvItems)
            {
                ws.Cells[r, 1] = $"{tvPrefix}{tvIdx++}";
                ws.Cells[r, 2] = "- " + item.TenChiPhi;
                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    string baseCell = item.CoSoTinh == CoSoTinhChiPhi.ChiPhiXayDung ? $"C{rowXD_Con}" : (item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi ? (hasTB ? $"C{rowTB}" : $"C{rowXD_Con}") : (hasTB ? $"(C{rowXD_Con}+C{rowTB})" : $"C{rowXD_Con}"));
                    string tlStr = (item.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string hsStr = item.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    ws.Cells[r, 3].Formula = $"=ROUND({tlStr} * {baseCell} * {hsStr}, 0)";
                }
                else
                {
                    ws.Cells[r, 3] = (double)item.GiaTriTruocThue;
                }
                string vatStr = item.ThueSuatGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatStr}, 0)";
                ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
                SetCellSymbolWithSubscript(ws, r, 6, item.KyHieu);
                r++;
            }
            int endTV = r - 1;
            if (endTV >= startTV)
            {
                ws.Cells[rowTVGroup, 3].Formula = $"=SUM(C{startTV}:C{endTV})";
                ws.Cells[rowTVGroup, 4].Formula = $"=SUM(D{startTV}:D{endTV})";
                ws.Cells[rowTVGroup, 5].Formula = $"=ROUND(SUM(E{startTV}:E{endTV}), -3)";
            }
            else
            {
                ws.Cells[rowTVGroup, 3] = 0; ws.Cells[rowTVGroup, 4] = 0; ws.Cells[rowTVGroup, 5] = 0;
            }

            // Chi phí khác (Cơ sở tính: Chi phí xây dựng con C{rowXD_Con})
            int rowKGroup = r;
            ws.Cells[r, 1] = hasTB ? "'6" : "'5";
            ws.Cells[r, 2] = "Chi phí khác";
            SetCellSymbolWithSubscript(ws, r, 6, "GK");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            int startK = r;
            var kItems = model.Items.Where(x => x.IsActive && x.Nhom == NhomChiPhi.ChiPhiKhac && (hasTB || x.CoSoTinh != CoSoTinhChiPhi.ChiPhiThietBi)).ToList();
            int kIdx = 1;
            string kPrefix = hasTB ? "'6." : "'5.";
            foreach (var item in kItems)
            {
                ws.Cells[r, 1] = $"{kPrefix}{kIdx++}";
                ws.Cells[r, 2] = "- " + item.TenChiPhi;
                if (item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc)
                {
                    string baseCell = item.CoSoTinh == CoSoTinhChiPhi.ChiPhiXayDung ? $"C{rowXD_Con}" : (item.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi ? (hasTB ? $"C{rowTB}" : $"C{rowXD_Con}") : (hasTB ? $"(C{rowXD_Con}+C{rowTB})" : $"C{rowXD_Con}"));
                    string tlStr = (item.TyLePhanTram / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string hsStr = item.HeSoDieuChinh.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (item.MinValue.HasValue && item.MaxValue.HasValue)
                    {
                        string minStr = item.MinValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        string maxStr = item.MaxValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        ws.Cells[r, 3].Formula = $"=ROUND(MAX({minStr}, MIN({maxStr}, {tlStr} * {baseCell} * {hsStr})), 0)";
                    }
                    else if (item.MinValue.HasValue)
                    {
                        string minStr = item.MinValue.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        ws.Cells[r, 3].Formula = $"=ROUND(MAX({minStr}, {tlStr} * {baseCell} * {hsStr}), 0)";
                    }
                    else
                    {
                        ws.Cells[r, 3].Formula = $"=ROUND({tlStr} * {baseCell} * {hsStr}, 0)";
                    }
                }
                else
                {
                    ws.Cells[r, 3] = (double)item.GiaTriTruocThue;
                }
                string vatStr = item.ThueSuatGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ws.Cells[r, 4].Formula = $"=ROUND(C{r} * {vatStr}, 0)";
                ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
                SetCellSymbolWithSubscript(ws, r, 6, item.KyHieu);
                r++;
            }
            int endK = r - 1;
            if (endK >= startK)
            {
                ws.Cells[rowKGroup, 3].Formula = $"=SUM(C{startK}:C{endK})";
                ws.Cells[rowKGroup, 4].Formula = $"=SUM(D{startK}:D{endK})";
                ws.Cells[rowKGroup, 5].Formula = $"=ROUND(SUM(E{startK}:E{endK}), -3)";
            }
            else
            {
                ws.Cells[rowKGroup, 3] = 0; ws.Cells[rowKGroup, 4] = 0; ws.Cells[rowKGroup, 5] = 0;
            }

            // Chi phí dự phòng (G_DP) - TMĐT 10%
            int rowDPGroup = r;
            ws.Cells[r, 1] = hasTB ? "'7" : "'6";
            ws.Cells[r, 2] = "Chi phí dự phòng";
            SetCellSymbolWithSubscript(ws, r, 6, "GDP");
            ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]].Font.Bold = true;
            r++;

            var itemDP = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong) ?? new ChiPhiKinhPhiItem { TyLePhanTram = 10.0m };
            string dpTyLe = (itemDP.TyLePhanTram > 5.0m ? itemDP.TyLePhanTram / 100m : 0.10m).ToString(System.Globalization.CultureInfo.InvariantCulture);
            int rowDP1 = r;
            ws.Cells[r, 1] = hasTB ? "'7.1" : "'6.1";
            ws.Cells[r, 2] = "- Chi phí dự phòng yếu tố khối lượng phát sinh (Gdp1)";
            string dpBase = hasTB
                ? $"(C{rowBT} + C{rowXD_Tong} + C{rowTB} + C{rowQLDA} + C{rowTVGroup} + C{rowKGroup})"
                : $"(C{rowBT} + C{rowXD_Tong} + C{rowQLDA} + C{rowTVGroup} + C{rowKGroup})";
            ws.Cells[r, 3].Formula = $"=ROUND({dpTyLe} * {dpBase}, 0)";
            ws.Cells[r, 4].Formula = $"=ROUND(C{r}*0.1, 0)";
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gdp1");
            r++;

            int rowDP2 = r;
            ws.Cells[r, 1] = hasTB ? "'7.2" : "'6.2";
            ws.Cells[r, 2] = "- Chi phí dự phòng yếu tố trượt giá (Gdp2)";
            ws.Cells[r, 3] = 0;
            ws.Cells[r, 4] = 0;
            ws.Cells[r, 5].Formula = $"=C{r}+D{r}";
            SetCellSymbolWithSubscript(ws, r, 6, "Gdp2");
            r++;

            ws.Cells[rowDPGroup, 3].Formula = $"=C{rowDP1}+C{rowDP2}";
            ws.Cells[rowDPGroup, 4].Formula = $"=D{rowDP1}+D{rowDP2}";
            ws.Cells[rowDPGroup, 5].Formula = $"=ROUND(E{rowDP1}+E{rowDP2}, -3)";

            // Dòng Tổng mức đầu tư xây dựng
            int grandRow = r;
            ws.Cells[r, 2] = "TỔNG MỨC ĐẦU TƯ XÂY DỰNG";
            string grandSumC = hasTB
                ? $"=C{rowBT}+C{rowXD_Tong}+C{rowTB}+C{rowQLDA}+C{rowTVGroup}+C{rowKGroup}+C{rowDPGroup}"
                : $"=C{rowBT}+C{rowXD_Tong}+C{rowQLDA}+C{rowTVGroup}+C{rowKGroup}+C{rowDPGroup}";
            string grandSumD = hasTB
                ? $"=D{rowBT}+D{rowXD_Tong}+D{rowTB}+D{rowQLDA}+D{rowTVGroup}+D{rowKGroup}+D{rowDPGroup}"
                : $"=D{rowBT}+D{rowXD_Tong}+D{rowQLDA}+D{rowTVGroup}+D{rowKGroup}+D{rowDPGroup}";
            string grandSumE = hasTB
                ? $"=ROUND(E{rowBT}+E{rowXD_Tong}+E{rowTB}+E{rowQLDA}+E{rowTVGroup}+E{rowKGroup}+E{rowDPGroup}, -3)"
                : $"=ROUND(E{rowBT}+E{rowXD_Tong}+E{rowQLDA}+E{rowTVGroup}+E{rowKGroup}+E{rowDPGroup}, -3)";
            ws.Cells[r, 3].Formula = grandSumC;
            ws.Cells[r, 4].Formula = grandSumD;
            ws.Cells[r, 5].Formula = grandSumE;
            SetCellSymbolWithSubscript(ws, r, 6, "GTM");

            var grandRng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 6]];
            grandRng.Font.Bold = true;
            grandRng.Font.Size = 12;
            grandRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));

            // Dòng Bằng chữ (Merge A..F, căn giữa, in nghiêng)
            int textRow = grandRow + 1;
            try { ws.Calculate(); } catch { }
            object valGrandTM = ws.Cells[grandRow, 5]?.Value2;
            decimal tongSauThue = 0m;
            if (valGrandTM is double d) tongSauThue = (decimal)d;
            else if (valGrandTM is decimal dec) tongSauThue = dec;
            else if (valGrandTM != null)
            {
                if (decimal.TryParse(valGrandTM.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal pVal) ||
                    decimal.TryParse(valGrandTM.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out pVal))
                {
                    tongSauThue = pVal;
                }
            }
            if (tongSauThue <= 0)
            {
                tongSauThue = model.TongSauThue;
            }
            tongSauThue = Math.Round(tongSauThue / 1000m, 0, MidpointRounding.AwayFromZero) * 1000m;
            string chuTien = UIHelper.DocSoThanhChu(tongSauThue);
            var textRange = ws.Range[ws.Cells[textRow, 1], ws.Cells[textRow, 6]];
            textRange.Merge();
            textRange.Value2 = $"Bằng chữ: {chuTien}.";
            textRange.Font.Italic = true;
            textRange.Font.Bold = false;
            textRange.Font.Size = 12;
            textRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            textRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            DrawTableBorders(ws, 5, 1, textRow, 6);
            ws.Range[$"C7:E{grandRow}"].NumberFormat = "#,##0;-#,##0;\"-\"";
            ws.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[6].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();

            // Cố định dòng 6 (Freeze Panes) để luôn nhìn thấy tiêu đề khi cuộn dọc
            ApplyFreezePanes(ws, 6);

            // Yêu cầu 3: Di chuyển sheet TongMucDauTu nằm ngay phía trước sheet TH_ChiPhiXD
            try
            {
                Worksheet wsTarget = GetSheetSafe(wb, "TH_ChiPhiXD");
                if (wsTarget == null)
                {
                    foreach (Worksheet sh in wb.Sheets)
                    {
                        if (sh.Name.StartsWith("DuToan")) { wsTarget = sh; break; }
                    }
                }
                if (wsTarget != null)
                {
                    ws.Move(Before: wsTarget);
                }
                else if (wb.Sheets.Count > 1)
                {
                    ws.Move(Before: wb.Sheets[1]);
                }
            }
            catch { }
        }

        private int FindDuToanRow(Worksheet wsDuToan, DongDuToan ct, HashSet<int> usedRows)
        {
            if (wsDuToan == null || ct == null) return -1;

            // 1. Kiểm tra ct.STT trước nếu chưa dùng và Mã hiệu khớp
            if (ct.STT >= 6 && !usedRows.Contains(ct.STT))
            {
                string mh = wsDuToan.Cells[ct.STT, 2]?.Value2?.ToString()?.Trim() ?? "";
                if (string.Equals(mh, ct.MaHieu, StringComparison.OrdinalIgnoreCase))
                {
                    usedRows.Add(ct.STT);
                    return ct.STT;
                }
            }

            // 2. Quét tuần tự từ dòng 6
            for (int rScan = 6; rScan <= 10000; rScan++)
            {
                if (usedRows.Contains(rScan)) continue;
                string ten = wsDuToan.Cells[rScan, 3]?.Value2?.ToString()?.Trim() ?? "";
                if (ten.Equals("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || ten.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                    break;
                string mh = wsDuToan.Cells[rScan, 2]?.Value2?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(mh) && string.IsNullOrEmpty(ten))
                    break;

                if (string.Equals(mh, ct.MaHieu, StringComparison.OrdinalIgnoreCase))
                {
                    usedRows.Add(rScan);
                    ct.STT = rScan;
                    return rScan;
                }
            }

            return -1;
        }

        public void XuatPhanTichDonGia(Workbook wb, DuToan duToan, Worksheet wsDuToan)
        {
            var ws = CreateOrGetSheet(wb, "PhanTich_DonGia");
            SetupHeader(ws, "BẢNG PHÂN TÍCH ĐƠN GIÁ CHI TIẾT", 7);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã hiệu";
            ws.Cells[3, 3] = "Tên công tác / Danh mục vật tư";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Định mức";
            ws.Cells[3, 6] = "Đơn giá (đồng)";
            ws.Cells[3, 7] = "Thành tiền (đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 7]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            int r = 4;
            int stt = 1;

            if (wsDuToan == null)
            {
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase))
                    {
                        wsDuToan = sheet;
                        break;
                    }
                }
            }

            var usedDuToanRows = new HashSet<int>();

            foreach (var hm in duToan.DanhSachHangMuc)
            {
                foreach (var ct in hm.DanhSachCongTac)
                {
                    // Dòng Công tác
                    ws.Cells[r, 1] = stt++;
                    ws.Cells[r, 2] = ct.MaHieu;
                    ws.Cells[r, 3] = ct.TenCongTac;
                    ws.Cells[r, 4] = ct.DonVi;
                    ws.Range[ws.Cells[r, 1], ws.Cells[r, 7]].Font.Bold = true;
                    ws.Range[ws.Cells[r, 1], ws.Cells[r, 7]].Interior.Color = ColorTranslator.ToOle(Color.LightGoldenrodYellow);
                    r++;

                    List<int> subTotalRows = new List<int>();
                    int? rowCongVL = null;
                    int? rowCongNC = null;
                    int? rowCongMay = null;

                    // Bảng Hao Phí
                    if (ct.DanhSachHaoPhi != null && ct.DanhSachHaoPhi.Any())
                    {
                        var vlHaoPhi = ct.DanhSachHaoPhi.Where(x => x.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.VL).ToList();
                        if (vlHaoPhi.Any())
                        {
                            ws.Cells[r, 3] = "Vật liệu";
                            var rngVL = (Range)ws.Cells[r, 3];
                            rngVL.Font.Bold = true;
                            rngVL.Font.Italic = true;
                            r++;
                            foreach (var hp in vlHaoPhi)
                            {
                                ws.Cells[r, 2] = hp.MaHieuHP;
                                ws.Cells[r, 3] = hp.TenHaoPhi;
                                ws.Cells[r, 4] = hp.DonVi;
                                ws.Cells[r, 5] = (double)(hp.DinhMuc * hp.HeSo);
                                
                                if (hp.DonVi != "%")
                                {
                                    ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_VatLieu'!$B:$I, 8, 0), 0)";
                                    ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                                }
                                else
                                {
                                    ws.Cells[r, 6].Formula = $"=SUM(G{r - vlHaoPhi.IndexOf(hp)}:G{r - 1})";
                                    ws.Cells[r, 7].Formula = $"=E{r}*F{r}/100";
                                }
                                r++;
                            }
                            ws.Cells[r, 3] = "Cộng chi phí Vật liệu";
                            ws.Cells[r, 7].Formula = $"=SUM(G{r - vlHaoPhi.Count}:G{r - 1})";
                            ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                            subTotalRows.Add(r);
                            rowCongVL = r;
                            r++;
                        }

                        var ncHaoPhi = ct.DanhSachHaoPhi.Where(x => x.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.NC).ToList();
                        if (ncHaoPhi.Any())
                        {
                            ws.Cells[r, 3] = "Nhân công";
                            var rngNC = (Range)ws.Cells[r, 3];
                            rngNC.Font.Bold = true;
                            rngNC.Font.Italic = true;
                            r++;
                            foreach (var hp in ncHaoPhi)
                            {
                                ws.Cells[r, 2] = hp.MaHieuHP;
                                ws.Cells[r, 3] = hp.TenHaoPhi;
                                ws.Cells[r, 4] = hp.DonVi;
                                ws.Cells[r, 5] = (double)(hp.DinhMuc * hp.HeSo);
                                ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_NhanCong'!$B:$E, 4, 0), 0)";
                                ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                                r++;
                            }
                            ws.Cells[r, 3] = "Cộng chi phí Nhân công";
                            ws.Cells[r, 7].Formula = $"=SUM(G{r - ncHaoPhi.Count}:G{r - 1})";
                            ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                            subTotalRows.Add(r);
                            rowCongNC = r;
                            r++;
                        }

                        var mayHaoPhi = ct.DanhSachHaoPhi.Where(x => x.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.MAY).ToList();
                        if (mayHaoPhi.Any())
                        {
                            ws.Cells[r, 3] = "Máy thi công";
                            var rngMay = (Range)ws.Cells[r, 3];
                            rngMay.Font.Bold = true;
                            rngMay.Font.Italic = true;
                            r++;
                            foreach (var hp in mayHaoPhi)
                            {
                                ws.Cells[r, 2] = hp.MaHieuHP;
                                ws.Cells[r, 3] = hp.TenHaoPhi;
                                ws.Cells[r, 4] = hp.DonVi;
                                ws.Cells[r, 5] = (double)(hp.DinhMuc * hp.HeSo);
                                
                                if (hp.DonVi != "%")
                                {
                                    ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_CaMay'!$B:$E, 4, 0), 0)";
                                    ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                                }
                                else
                                {
                                    ws.Cells[r, 6].Formula = $"=SUM(G{r - mayHaoPhi.IndexOf(hp)}:G{r - 1})";
                                    ws.Cells[r, 7].Formula = $"=E{r}*F{r}/100";
                                }
                                r++;
                            }
                            ws.Cells[r, 3] = "Cộng chi phí Máy thi công";
                            ws.Cells[r, 7].Formula = $"=SUM(G{r - mayHaoPhi.Count}:G{r - 1})";
                            ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                            subTotalRows.Add(r);
                            rowCongMay = r;
                            r++;
                        }
                    }
                    
                    ws.Cells[r, 3] = "ĐƠN GIÁ TỔNG HỢP";
                    if (subTotalRows.Count > 0) {
                        ws.Cells[r, 7].Formula = "=" + string.Join("+", subTotalRows.Select(row => $"G{row}"));
                    } else {
                        ws.Cells[r, 7] = 0;
                    }
                    ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                    r++;

                    // Gán công thức link từ PhanTich_DonGia sang sheet DuToan
                    if (wsDuToan != null)
                    {
                        int rDu = FindDuToanRow(wsDuToan, ct, usedDuToanRows);
                        if (rDu >= 6)
                        {
                            if (rowCongVL.HasValue)
                                wsDuToan.Cells[rDu, 6].Formula = $"='PhanTich_DonGia'!G{rowCongVL.Value}";
                            else if (ct.DonGiaVL == 0)
                                wsDuToan.Cells[rDu, 6].Value2 = 0;

                            if (rowCongNC.HasValue)
                                wsDuToan.Cells[rDu, 7].Formula = $"='PhanTich_DonGia'!G{rowCongNC.Value}";
                            else if (ct.DonGiaNC == 0)
                                wsDuToan.Cells[rDu, 7].Value2 = 0;

                            if (rowCongMay.HasValue)
                                wsDuToan.Cells[rDu, 8].Formula = $"='PhanTich_DonGia'!G{rowCongMay.Value}";
                            else if (ct.DonGiaMay == 0)
                                wsDuToan.Cells[rDu, 8].Value2 = 0;

                            wsDuToan.Cells[rDu, 9].Formula = $"=ROUND(E{rDu}*F{rDu}, 0)";
                            wsDuToan.Cells[rDu, 10].Formula = $"=ROUND(E{rDu}*G{rDu}, 0)";
                            wsDuToan.Cells[rDu, 11].Formula = $"=ROUND(E{rDu}*H{rDu}, 0)";
                        }
                    }
                }
            }

            DrawTableBorders(ws, 3, 1, r - 1, 7);
            ws.Range["F:F"].NumberFormat = "#,##0";
            ws.Range["G:G"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
            ApplyFreezePanes(ws, 3);
        }

        /// <summary>
        /// Xuất bảng chiết tính cước vận chuyển vật liệu bằng ô tô (TT 38/2026/TT-BXD) và vận chuyển bộ (AM.21000)
        /// Các công thức liên kết 100% với TH_CaMay, TH_NhanCong và link ngược lên bảng tổng hợp cước
        /// </summary>
        public void XuatChietTinhCuocVC(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "ChietTinh_CuocVC");
            SetupHeader(ws, "BẢNG CHIẾT TÍNH CƯỚC VẬN CHUYỂN VẬT LIỆU", 13);

            // Phụ đề
            ws.Cells[2, 1] = "Áp dụng định mức vận chuyển Chương XII - Thông tư số 12/2021/TT-BXD và Thông tư số 38/2026/TT-BXD";
            var subTitleRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, 13]];
            subTitleRange.Merge();
            subTitleRange.Font.Italic = true;
            subTitleRange.Font.Size = 10;
            subTitleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // =========================================================================
            // I. BẢNG TỔNG HỢP CƯỚC VẬN CHUYỂN CÁC LOẠI VẬT LIỆU
            // =========================================================================
            ws.Cells[4, 1] = "I. BẢNG TỔNG HỢP CƯỚC VẬN CHUYỂN VẬT LIỆU";
            var sec1Rng = (Range)ws.Cells[4, 1];
            sec1Rng.Font.Bold = true;
            sec1Rng.Font.Size = 11;
            sec1Rng.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 70, 140));

            string[] headers1 = new string[]
            {
                "STT", "Mã vật liệu", "Tên vật liệu", "Đơn vị", 
                "Phương tiện VC ô tô", "Tổng cự ly ô tô (km)", "Cước ô tô (đồng/ĐVT)", 
                "Cự ly bộ (m)", "Cước bộ (đồng/ĐVT)", "Tổng cước VC (đồng/ĐVT)"
            };

            for (int c = 0; c < headers1.Length; c++)
            {
                ws.Cells[5, c + 1] = headers1[c];
            }
            var h1Rng = ws.Range[ws.Cells[5, 1], ws.Cells[5, 10]];
            h1Rng.Font.Bold = true;
            h1Rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            h1Rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            h1Rng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(235, 241, 247));

            var danhSachVL = duToan.BangTongHop.DanhSachVatLieu ?? new List<VatLieuHienTruong>();
            var matRowInTable1 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            int r = 6;
            int stt1 = 1;
            foreach (var vl in danhSachVL)
            {
                matRowInTable1[vl.MaVatTu] = r;
                ws.Cells[r, 1] = stt1++;
                ws.Cells[r, 2] = vl.MaVatTu;
                ws.Cells[r, 3] = vl.TenVatTu;
                ws.Cells[r, 4] = vl.DonVi;
                ws.Cells[r, 5] = "-";
                ws.Cells[r, 6] = 0; // Cự ly ô tô (sẽ link sau)
                ws.Cells[r, 7] = (double)vl.CuocVCOTo; // Cước ô tô (sẽ link sau)
                ws.Cells[r, 8] = 0; // Cự ly bộ (sẽ link sau)
                ws.Cells[r, 9] = (double)vl.CuocVCBo;  // Cước bộ (sẽ link sau)
                ws.Cells[r, 10].Formula = $"=G{r}+I{r}";
                r++;
            }

            int endTable1Row = r - 1;
            if (endTable1Row >= 6)
            {
                DrawTableBorders(ws, 5, 1, endTable1Row, 10);
                ws.Range[$"F6:F{endTable1Row}"].NumberFormat = "#,##0.000";
                ws.Range[$"G6:G{endTable1Row}"].NumberFormat = "#,##0";
                ws.Range[$"H6:H{endTable1Row}"].NumberFormat = "#,##0";
                ws.Range[$"I6:J{endTable1Row}"].NumberFormat = "#,##0";
            }
            ws.Columns[1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            r += 2; // Spacing

            // =========================================================================
            // II. CHI TIẾT CHIẾT TÍNH CƯỚC VẬN CHUYỂN BẰNG Ô TÔ (THÔNG TƯ 38/2026/TT-BXD)
            // =========================================================================
            ws.Cells[r, 1] = "II. CHI TIẾT CHIẾT TÍNH CƯỚC VẬN CHUYỂN BẰNG Ô TÔ (Chương XII - Thông tư số 38/2026/TT-BXD)";
            var sec2Rng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 13]];
            sec2Rng.Merge();
            sec2Rng.Font.Bold = true;
            sec2Rng.Font.Size = 12;
            sec2Rng.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
            sec2Rng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));
            r += 2;

            int matIdx = 1;
            foreach (var vl in danhSachVL)
            {
                var cfg = VanChuyenStorage.GetConfigOTo(vl.TenVatTu);
                DinhMucVCOToItem? dm = null;

                if (cfg != null && !string.IsNullOrEmpty(cfg.MaDinhMuc))
                {
                    dm = DinhMucVanChuyenDatabase.DanhSachOTo.FirstOrDefault(x => x.MaHieu == cfg.MaDinhMuc);
                }
                if (dm == null)
                {
                    dm = DinhMucVanChuyenDatabase.NhanDienOTo(vl.TenVatTu);
                }

                // Chỉ chiết tính ô tô cho các vật liệu người dùng ĐÃ cấu hình hoặc CÓ cước ô tô > 0
                bool coCuocOTo = vl.CuocVCOTo > 0 || (cfg != null && (cfg.TongCuLyKm > 0 || cfg.CuocVCOTo > 0));
                if (!coCuocOTo)
                {
                    continue;
                }

                if (dm == null)
                {
                    dm = DinhMucVanChuyenDatabase.DanhSachOTo.First(x => x.MaHieu == "AM.2411");
                }

                List<CungDuongVanChuyen> cungDuongs = cfg?.CungDuongs?.Count > 0 
                    ? cfg.CungDuongs 
                    : new List<CungDuongVanChuyen>
                    {
                        new() { DiemDau = "Nơi cung cấp", DiemCuoi = "Công trình", TenDoanDuong = "Tuyến chính", CuLyKm = (cfg?.TongCuLyKm > 0 ? cfg.TongCuLyKm : 10m), LoaiDuong = 3 }
                    };

                decimal donGiaCaMay = cfg != null && cfg.DonGiaCaMay > 0 ? cfg.DonGiaCaMay : 1040324m;

                // Tên mục vật liệu
                ws.Cells[r, 1] = $"{matIdx++}. Vật liệu: {vl.TenVatTu} (Mã hiệu: {vl.MaVatTu}) - Đơn vị tính: {vl.DonVi}";
                var matTitleRng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 13]];
                matTitleRng.Merge();
                matTitleRng.Font.Bold = true;
                matTitleRng.Font.Size = 10.5;
                matTitleRng.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 102, 204));
                matTitleRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 248, 255));
                r++;

                // Dòng thông tin phương tiện & giá ca xe
                ws.Cells[r, 2] = "Phương tiện:";
                ws.Cells[r, 3] = $"{dm.MaHieu} - {dm.TenCongTac}";
                ws.Cells[r, 5] = "Mã ca máy:";
                ws.Cells[r, 6] = dm.MaMay;
                ws.Cells[r, 8] = "Giá ca xe (đồng/ca):";
                ws.Cells[r, 9].Formula = $"=IFERROR(VLOOKUP(F{r}, 'TH_CaMay'!$B:$E, 4, 0), {donGiaCaMay.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                ws.Range[ws.Cells[r, 2], ws.Cells[r, 9]].Font.Bold = true;
                ws.Cells[r, 9].NumberFormat = "#,##0";
                int giaCaXeRow = r;
                r++;

                // Cập nhật phương tiện vào Table 1
                if (matRowInTable1.TryGetValue(vl.MaVatTu, out int t1Row))
                {
                    ws.Cells[t1Row, 5] = dm.TenMay;
                }

                // Tiêu đề bảng đoạn đường
                string[] segHeaders = new string[]
                {
                    "STT", "Điểm đầu", "Điểm cuối", "Tên tuyến đường / đoạn đường",
                    "Cự ly L (km)", "Loại đường", "Hệ số kđ", "Km đầu", "Km cuối",
                    "Quy đổi Nấc 1 (<=1km)", "Quy đổi Nấc 2 (1-10km)", "Quy đổi Nấc 3 (10-60km)", "Quy đổi Nấc 4 (>60km)"
                };

                for (int sc = 0; sc < segHeaders.Length; sc++)
                {
                    ws.Cells[r, sc + 1] = segHeaders[sc];
                }
                var segHRange = ws.Range[ws.Cells[r, 1], ws.Cells[r, 13]];
                segHRange.Font.Bold = true;
                segHRange.Font.Size = 9;
                segHRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                segHRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                segHRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
                int segHeaderRow = r;
                r++;

                int startSegRow = r;
                int segIdx = 1;
                foreach (var cd in cungDuongs)
                {
                    ws.Cells[r, 1] = segIdx++;
                    ws.Cells[r, 2] = cd.DiemDau;
                    ws.Cells[r, 3] = cd.DiemCuoi;
                    ws.Cells[r, 4] = cd.TenDoanDuong;
                    ws.Cells[r, 5] = (double)cd.CuLyKm;
                    ws.Cells[r, 6] = $"Loại {cd.LoaiDuong}";
                    ws.Cells[r, 7] = (double)cd.HeSoK;

                    // Km đầu (H) và Km cuối (I)
                    if (r == startSegRow)
                    {
                        ws.Cells[r, 8] = 0;
                    }
                    else
                    {
                        ws.Cells[r, 8].Formula = $"=I{r - 1}";
                    }
                    ws.Cells[r, 9].Formula = $"=H{r}+E{r}";

                    // 4 nấc cự ly quy đổi
                    ws.Cells[r, 10].Formula = $"=MAX(0, MIN(I{r}, 1) - MAX(H{r}, 0)) * G{r}";
                    ws.Cells[r, 11].Formula = $"=MAX(0, MIN(I{r}, 10) - MAX(H{r}, 1)) * G{r}";
                    ws.Cells[r, 12].Formula = $"=MAX(0, MIN(I{r}, 60) - MAX(H{r}, 10)) * G{r}";
                    ws.Cells[r, 13].Formula = $"=MAX(0, I{r} - MAX(H{r}, 60)) * G{r}";

                    r++;
                }
                int endSegRow = r - 1;

                // Dòng Tổng cộng cự ly và km quy đổi
                ws.Cells[r, 4] = "Cộng cự ly quy đổi:";
                ws.Cells[r, 5].Formula = $"=SUM(E{startSegRow}:E{endSegRow})";
                ws.Cells[r, 10].Formula = $"=SUM(J{startSegRow}:J{endSegRow})";
                ws.Cells[r, 11].Formula = $"=SUM(K{startSegRow}:K{endSegRow})";
                ws.Cells[r, 12].Formula = $"=SUM(L{startSegRow}:L{endSegRow})";
                ws.Cells[r, 13].Formula = $"=SUM(M{startSegRow}:M{endSegRow})";
                ws.Range[ws.Cells[r, 4], ws.Cells[r, 13]].Font.Bold = true;
                int sumQdRow = r;
                int tongCuLyRow = r;
                r++;

                // Dòng Định mức hao phí ca xe
                ws.Cells[r, 4] = $"Định mức hao phí ca xe ({dm.DonViDinhMuc}/km):";
                ws.Cells[r, 10] = (double)dm.Dm1;
                ws.Cells[r, 11] = (double)dm.Dm2;
                ws.Cells[r, 12] = (double)dm.Dm3;
                ws.Cells[r, 13] = (double)(dm.Dm3 * 0.95m);
                int dmRow = r;
                r++;

                // Dòng Hao phí ca xe theo từng nấc
                ws.Cells[r, 4] = "Hao phí ca xe theo nấc (ca):";
                ws.Cells[r, 10].Formula = $"=J{sumQdRow}*J{dmRow}";
                ws.Cells[r, 11].Formula = $"=K{sumQdRow}*K{dmRow}";
                ws.Cells[r, 12].Formula = $"=L{sumQdRow}*L{dmRow}";
                ws.Cells[r, 13].Formula = $"=M{sumQdRow}*M{dmRow}";
                int hpRow = r;
                r++;

                // Bảng kẻ viền cho phần chi tiết đoạn đường
                DrawTableBorders(ws, segHeaderRow, 1, hpRow, 13);
                ws.Range[$"E{startSegRow}:E{sumQdRow}"].NumberFormat = "#,##0.000";
                ws.Range[$"G{startSegRow}:G{endSegRow}"].NumberFormat = "0.00";
                ws.Range[$"H{startSegRow}:M{hpRow}"].NumberFormat = "#,##0.0000";

                // Tổng hợp kết quả cước ô tô
                ws.Cells[r, 4] = $"Tổng hao phí ca xe cho 1 {dm.DonViDinhMuc} (ca):";
                ws.Cells[r, 5].Formula = $"=SUM(J{hpRow}:M{hpRow})";
                ws.Cells[r, 5].Font.Bold = true;
                ws.Cells[r, 5].NumberFormat = "#,##0.0000";
                int tongHpCaRow = r;
                r++;

                ws.Cells[r, 4] = $"Chi phí vận chuyển ô tô cho 1 {dm.DonViDinhMuc} (đồng):";
                ws.Cells[r, 5].Formula = $"=E{tongHpCaRow}*I{giaCaXeRow}";
                ws.Cells[r, 5].Font.Bold = true;
                ws.Cells[r, 5].NumberFormat = "#,##0";
                int cuocDvdRow = r;
                r++;

                decimal heSoQd = DinhMucVanChuyenDatabase.TinhHeSoQuyDoiOTo(vl.DonVi, dm.DonViDinhMuc);
                ws.Cells[r, 4] = $"Hệ số quy đổi về 1 {vl.DonVi}:";
                ws.Cells[r, 5] = (double)heSoQd;
                ws.Cells[r, 5].NumberFormat = "#,##0.0000";
                int heSoQdRow = r;
                r++;

                ws.Cells[r, 4] = $"CƯỚC VẬN CHUYỂN Ô TÔ / 1 {vl.DonVi} (đồng):";
                ws.Cells[r, 5].Formula = $"=E{cuocDvdRow}*E{heSoQdRow}";
                ws.Range[ws.Cells[r, 4], ws.Cells[r, 5]].Font.Bold = true;
                ws.Range[ws.Cells[r, 4], ws.Cells[r, 5]].Font.Size = 10.5;
                ws.Range[ws.Cells[r, 4], ws.Cells[r, 5]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));
                ws.Cells[r, 5].NumberFormat = "#,##0";
                int cuocOToFinalRow = r;
                r++;

                // LINK VÀO TABLE 1 Ở TRÊN:
                if (matRowInTable1.TryGetValue(vl.MaVatTu, out int targetRow))
                {
                    ws.Cells[targetRow, 6].Formula = $"=E{tongCuLyRow}";
                    ws.Cells[targetRow, 7].Formula = $"=E{cuocOToFinalRow}";
                }

                r += 2; // Spacing giữa các vật liệu
            }

            // =========================================================================
            // III. CHI TIẾT CHIẾT TÍNH CƯỚC VẬN CHUYỂN BỘ (THỦ CÔNG)
            // =========================================================================
            var materialsWithBo = danhSachVL.Where(v => 
            {
                var cfgBo = VanChuyenStorage.GetConfigBo(v.TenVatTu);
                return (cfgBo != null && cfgBo.CuLyMet > 0) || v.CuocVCBo > 0;
            }).ToList();

            if (materialsWithBo.Count > 0)
            {
                ws.Cells[r, 1] = "III. CHI TIẾT CHIẾT TÍNH CƯỚC VẬN CHUYỂN BỘ (THỦ CÔNG - AM.21000)";
                var sec3Rng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 10]];
                sec3Rng.Merge();
                sec3Rng.Font.Bold = true;
                sec3Rng.Font.Size = 12;
                sec3Rng.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 51, 102));
                sec3Rng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(217, 225, 242));
                r += 2;

                int boIdx = 1;
                // Tìm nhân công xây dựng nhóm 1 trong DanhSachNhanCong để lấy mã hiệu và đơn giá thực tế
                var ncNhom1 = duToan.BangTongHop?.DanhSachNhanCong?.FirstOrDefault(n => 
                    n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && 
                    (n.TenVatTu.ToLower().Contains("nhóm 1") || n.TenVatTu.ToLower().Contains("nhóm i") || n.MaVatTu.EndsWith(".01")));
                if (ncNhom1 == null)
                {
                    ncNhom1 = duToan.BangTongHop?.DanhSachNhanCong?.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung);
                }
                if (ncNhom1 == null)
                {
                    ncNhom1 = duToan.BangTongHop?.DanhSachNhanCong?.FirstOrDefault();
                }

                decimal defaultGiaNC = ncNhom1 != null && ncNhom1.GiaHienTruong > 0 ? ncNhom1.GiaHienTruong : 254498m;
                string defaultMaNC = ncNhom1 != null ? ncNhom1.MaVatTu : "NC_XD_1";

                foreach (var vl in materialsWithBo)
                {
                    var cfgBo = VanChuyenStorage.GetConfigBo(vl.TenVatTu);
                    var dmBo = DinhMucVanChuyenDatabase.NhanDienBo(vl.TenVatTu) ?? DinhMucVanChuyenDatabase.DanhSachBo.First();

                    decimal cuLyMet = cfgBo != null && cfgBo.CuLyMet > 0 ? cfgBo.CuLyMet : 30m;
                    decimal heSoDiaHinh = cfgBo != null && cfgBo.HeSoDiaHinh > 0 ? cfgBo.HeSoDiaHinh : 1.0m;
                    int soTang = cfgBo != null && cfgBo.SoTang > 0 ? cfgBo.SoTang : 1;

                    decimal giaNhanCong = (cfgBo != null && cfgBo.DonGiaNhanCong > 0) ? cfgBo.DonGiaNhanCong : defaultGiaNC;
                    string maNC = (cfgBo != null && !string.IsNullOrEmpty(cfgBo.MaNhanCong)) ? cfgBo.MaNhanCong : defaultMaNC;

                    ws.Cells[r, 1] = $"{boIdx++}. Vận chuyển bộ: {vl.TenVatTu} (Mã hiệu: {vl.MaVatTu}) - ĐVT: {vl.DonVi}";
                    var boTitleRng = ws.Range[ws.Cells[r, 1], ws.Cells[r, 10]];
                    boTitleRng.Merge();
                    boTitleRng.Font.Bold = true;
                    boTitleRng.Font.Size = 10.5;
                    boTitleRng.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 102, 204));
                    boTitleRng.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 248, 255));
                    r++;

                    // Bảng tham số
                    int startBoTable = r;
                    ws.Cells[r, 2] = "Mã định mức:"; ws.Cells[r, 3] = dmBo.MaHieu;
                    ws.Cells[r, 4] = "Công tác:"; ws.Cells[r, 5] = dmBo.TenCongTac;
                    r++;

                    ws.Cells[r, 2] = "Cự ly vận chuyển bộ L (m):"; ws.Cells[r, 3] = (double)cuLyMet;
                    int cuLyBoRow = r;
                    ws.Cells[r, 4] = "Hệ số địa hình kđh:"; ws.Cells[r, 5] = (double)heSoDiaHinh;
                    int heSoDhRow = r;
                    ws.Cells[r, 6] = "Số tầng cao:"; ws.Cells[r, 7] = soTang;
                    int soTangRow = r;
                    r++;

                    ws.Cells[r, 2] = $"Đm 10m đầu (công/{dmBo.DonViDinhMuc}):"; ws.Cells[r, 3] = (double)dmBo.Dm10m;
                    int dm10Row = r;
                    ws.Cells[r, 4] = $"Đm 10m tiếp theo (công/{dmBo.DonViDinhMuc}):"; ws.Cells[r, 5] = (double)dmBo.DmTiepTheo;
                    int dmTiepRow = r;
                    r++;

                    // Hao phí NC (công/ĐVDM)
                    ws.Cells[r, 2] = $"Hao phí nhân công cho 1 {dmBo.DonViDinhMuc} (công):";
                    ws.Cells[r, 3].Formula = $"=(C{dm10Row} + IF(C{cuLyBoRow}>10, (MIN(C{cuLyBoRow}, 300)-10)/10 * E{dmTiepRow}, 0)) * E{heSoDhRow} * (1.1^(G{soTangRow}-1))";
                    ws.Cells[r, 3].NumberFormat = "#,##0.0000";
                    ws.Cells[r, 3].Font.Bold = true;
                    int hpNcRow = r;
                    r++;

                    // Đơn giá nhân công (link từ TH_NhanCong)
                    ws.Cells[r, 2] = "Đơn giá nhân công (đồng/công):";
                    ws.Cells[r, 3].Formula = $"=IFERROR(VLOOKUP(\"{maNC}\", 'TH_NhanCong'!$B:$E, 4, 0), {giaNhanCong.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                    ws.Cells[r, 3].NumberFormat = "#,##0";
                    int giaNcRow = r;
                    r++;

                    // Thành tiền cho 1 ĐVDM
                    ws.Cells[r, 2] = $"Chi phí cước bộ cho 1 {dmBo.DonViDinhMuc} (đồng):";
                    ws.Cells[r, 3].Formula = $"=C{hpNcRow}*C{giaNcRow}";
                    ws.Cells[r, 3].NumberFormat = "#,##0";
                    int cuocBoDvdRow = r;
                    r++;

                    decimal heSoQdBo = DinhMucVanChuyenDatabase.TinhHeSoQuyDoiBo(vl.DonVi, dmBo.DonViDinhMuc);
                    ws.Cells[r, 2] = $"Hệ số quy đổi về 1 {vl.DonVi}:";
                    ws.Cells[r, 3] = (double)heSoQdBo;
                    ws.Cells[r, 3].NumberFormat = "#,##0.0000";
                    int heSoBoRow = r;
                    r++;

                    ws.Cells[r, 2] = $"CƯỚC VẬN CHUYỂN BỘ / 1 {vl.DonVi} (đồng):";
                    ws.Cells[r, 3].Formula = $"=C{cuocBoDvdRow}*C{heSoBoRow}";
                    ws.Range[ws.Cells[r, 2], ws.Cells[r, 3]].Font.Bold = true;
                    ws.Range[ws.Cells[r, 2], ws.Cells[r, 3]].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204));
                    ws.Cells[r, 3].NumberFormat = "#,##0";
                    int cuocBoFinalRow = r;
                    r++;

                    DrawTableBorders(ws, startBoTable, 2, cuocBoFinalRow, 7);

                    // LINK VÀO TABLE 1 Ở TRÊN:
                    if (matRowInTable1.TryGetValue(vl.MaVatTu, out int t1BoRow))
                    {
                        ws.Cells[t1BoRow, 8].Formula = $"=C{cuLyBoRow}";
                        ws.Cells[t1BoRow, 9].Formula = $"=C{cuocBoFinalRow}";
                    }

                    r += 2;
                }
            }

            ws.Columns.AutoFit();
        }

        public void XuatBangTongHopVatLieu(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_VatLieu");
            SetupHeader(ws, "BẢNG TỔNG HỢP VẬT LIỆU", 9);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã vật liệu";
            ws.Cells[3, 3] = "Tên vật liệu";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Giá gốc\n(đồng)";
            ws.Cells[3, 6] = "Chi phí bốc xếp\n(đồng)";
            ws.Cells[3, 7] = "Cước VC ô tô\n(đồng)";
            ws.Cells[3, 8] = "Cước VC bộ\n(đồng)";
            ws.Cells[3, 9] = "Giá hiện trường\n(đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 9]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignTop;
            headerRange.WrapText = true;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ws.Rows[3].RowHeight = 28;

            int r = 4;
            int stt = 1;
            foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = vl.MaVatTu;
                ws.Cells[r, 3] = vl.TenVatTu;
                ws.Cells[r, 4] = vl.DonVi;
                ws.Cells[r, 5] = (double)vl.GiaGoc;
                ws.Cells[r, 6] = (double)vl.ChiPhiBocXep;
                ws.Cells[r, 7].Formula = $"=IFERROR(VLOOKUP(B{r}, 'ChietTinh_CuocVC'!$B:$J, 6, 0), {((double)vl.CuocVCOTo).ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                ws.Cells[r, 8].Formula = $"=IFERROR(VLOOKUP(B{r}, 'ChietTinh_CuocVC'!$B:$J, 8, 0), {((double)vl.CuocVCBo).ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                ws.Cells[r, 9].Formula = $"=E{r}+F{r}+G{r}+H{r}";
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 9);
            ws.Range["E:I"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
            ApplyFreezePanes(ws, 3);
        }

        public void XuatBangTongHopNhanCong(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_NhanCong");
            SetupHeader(ws, "BẢNG TỔNG HỢP NHÂN CÔNG", 5);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã nhân công";
            ws.Cells[3, 3] = "Tên nhân công";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Giá nhân công\n(đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 5]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignTop;
            headerRange.WrapText = true;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ws.Rows[3].RowHeight = 28;

            int r = 4;
            int stt = 1;
            foreach (var nc in duToan.BangTongHop.DanhSachNhanCong)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = nc.MaVatTu;
                ws.Cells[r, 3] = nc.TenVatTu;
                ws.Cells[r, 4] = nc.DonVi;
                ws.Cells[r, 5] = (double)nc.GiaHienTruong;
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 5);
            ws.Range["E:E"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
            ApplyFreezePanes(ws, 3);
        }

        public void XuatBangTongHopCaMay(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_CaMay");
            SetupHeader(ws, "BẢNG TỔNG HỢP MÁY THI CÔNG", 5);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã ca máy";
            ws.Cells[3, 3] = "Tên loại máy";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Giá ca máy\n(đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 5]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignTop;
            headerRange.WrapText = true;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            ws.Rows[3].RowHeight = 28;

            int r = 4;
            int stt = 1;
            foreach (var m in duToan.BangTongHop.DanhSachMay)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = m.MaVatTu;
                ws.Cells[r, 3] = m.TenVatTu;
                ws.Cells[r, 4] = m.DonVi;
                ws.Cells[r, 5] = (double)m.GiaHienTruong;
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 5);
            ws.Range["E:E"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
            ApplyFreezePanes(ws, 3);
        }

        public void XuatBangHeSoDieuChinh(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "HeSo_DieuChinh");
            SetupHeader(ws, "BẢNG TỔNG HỢP HỆ SỐ VÀ TỶ LỆ", 3);
            
            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Loại hệ số / Tỷ lệ";
            ws.Cells[3, 3] = "Giá trị (%)";

            var kq = duToan.ChiPhiXD;
            if (kq == null) return;
            
            ws.Cells[4, 1] = 1; ws.Cells[4, 2] = "Tỷ lệ Chi phí chung (CPC)"; ws.Cells[4, 3] = (double)kq.TiLeCPC;
            ws.Cells[5, 1] = 2; ws.Cells[5, 2] = "Tỷ lệ CP không xác định KL (TT)"; ws.Cells[5, 3] = (double)kq.TiLeTT;
            ws.Cells[6, 1] = 3; ws.Cells[6, 2] = "Lợi nhuận định mức (TNCTTT)"; ws.Cells[6, 3] = (double)kq.TiLeTNCTTT;
            ws.Cells[7, 1] = 4; ws.Cells[7, 2] = "Thuế GTGT"; ws.Cells[7, 3] = (double)kq.TiLeGTGT;
            ws.Cells[8, 1] = 5; ws.Cells[8, 2] = "Chi phí nhà tạm (LT)"; ws.Cells[8, 3] = (double)kq.TiLeNhaTam;

            DrawTableBorders(ws, 3, 1, 8, 3);
            ws.Columns.AutoFit();
        }

        /// <summary>
        /// Áp giá hiện trường vào dự toán: Xuất các bảng phụ thuộc (TH_NhanCong, TH_CaMay, ChietTinh_CuocVC, TH_VatLieu, PhanTich_DonGia)
        /// và liên kết công thức sống (=PhanTich_DonGia!G...) vào sheet DuToan.
        /// </summary>
        public void ApGiaVaLienKetDuToan(DuToan duToan)
        {
            var app = (Application)ExcelDnaUtil.Application;
            var wb = app?.ActiveWorkbook;
            if (wb == null) throw new Exception("Không có Workbook nào đang mở.");

            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;

            try
            {
                Worksheet wsDuToan = null;
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name.StartsWith("DuToan", StringComparison.OrdinalIgnoreCase))
                    {
                        wsDuToan = sheet;
                        break;
                    }
                }
                if (wsDuToan == null) wsDuToan = wb.ActiveSheet as Worksheet;

                // 1. Xuất/cập nhật các bảng thành phần theo thứ tự phụ thuộc
                XuatBangTongHopNhanCong(wb, duToan);
                XuatBangTongHopCaMay(wb, duToan);
                XuatChietTinhCuocVC(wb, duToan);
                XuatBangTongHopVatLieu(wb, duToan);

                // 2. Xuất bảng phân tích đơn giá chi tiết và link công thức vào sheet DuToan
                XuatPhanTichDonGia(wb, duToan, wsDuToan);

                // 3. Sắp xếp lại thứ tự sheet theo đúng chuẩn
                SapXepLaiThuTuCacSheet(wb, wsDuToan);

                // 4. Cập nhật dòng TỔNG CỘNG trên sheet DuToan nếu có
                if (wsDuToan != null)
                {
                    int maxR = 5;
                    foreach (var hm in duToan.DanhSachHangMuc)
                    {
                        foreach (var dong in hm.DanhSachCongTac)
                        {
                            int r = dong.STT;
                            if (r > maxR) maxR = r;
                        }
                    }
                    int totalRow = maxR + 1;
                    string cVal = wsDuToan.Cells[totalRow, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (cVal.Equals("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase) || cVal.Equals("CỘNG", StringComparison.OrdinalIgnoreCase))
                    {
                        wsDuToan.Cells[totalRow, 9].Formula = $"=SUM(I6:I{maxR})";
                        wsDuToan.Cells[totalRow, 10].Formula = $"=SUM(J6:J{maxR})";
                        wsDuToan.Cells[totalRow, 11].Formula = $"=SUM(K6:K{maxR})";
                    }

                    wsDuToan.Activate();
                }

                // 5. Kích hoạt tính toán lại toàn bộ công thức
                app.Calculation = XlCalculation.xlCalculationAutomatic;
                try { app.Calculate(); } catch { }
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Calculation = XlCalculation.xlCalculationAutomatic;
            }
        }
    }
}
