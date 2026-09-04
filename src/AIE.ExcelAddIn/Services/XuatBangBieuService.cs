using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using AIE.Core.Models;

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

                // Bảng 1: Bảng tổng hợp dự toán (TH_DuToan)
                XuatBangTongHopDuToan(wb, duToan);

                // Bảng 2: Bảng phân tích đơn giá chi tiết (PhanTich_DonGia)
                XuatPhanTichDonGia(wb, duToan, wsDuToan);

                // Sắp xếp lại thứ tự sheet theo chuẩn hồ sơ dự toán
                try
                {
                    var wsTHDuToan = wb.Sheets["TH_DuToan"] as Worksheet;
                    var wsPhanTich = wb.Sheets["PhanTich_DonGia"] as Worksheet;
                    var wsTHVL = wb.Sheets["TH_VatLieu"] as Worksheet;
                    var wsCuocVC = wb.Sheets["ChietTinh_CuocVC"] as Worksheet;
                    var wsTHNC = wb.Sheets["TH_NhanCong"] as Worksheet;
                    var wsTHMay = wb.Sheets["TH_CaMay"] as Worksheet;
                    var wsHeSo = wb.Sheets["HeSo_DieuChinh"] as Worksheet;

                    if (wsTHDuToan != null) wsTHDuToan.Move(Before: wb.Sheets[1]);
                    if (wsDuToan != null && wsTHDuToan != null) wsDuToan.Move(After: wsTHDuToan);
                    if (wsPhanTich != null && wsDuToan != null) wsPhanTich.Move(After: wsDuToan);
                    if (wsTHVL != null && wsPhanTich != null) wsTHVL.Move(After: wsPhanTich);
                    if (wsCuocVC != null && wsTHVL != null) wsCuocVC.Move(After: wsTHVL);
                    if (wsTHNC != null && wsCuocVC != null) wsTHNC.Move(After: wsCuocVC);
                    if (wsTHMay != null && wsTHNC != null) wsTHMay.Move(After: wsTHNC);
                    if (wsHeSo != null && wsTHMay != null) wsHeSo.Move(After: wsTHMay);
                }
                catch { }

                // Bảng 7: Bảng xác định hệ số (HeSo_DieuChinh)
                XuatBangHeSoDieuChinh(wb, duToan);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Calculation = XlCalculation.xlCalculationAutomatic;
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
        private void ReformatDuToanHeader(Worksheet ws, DuToan duToan)
        {
            try
            {
                // Kiểm tra xem đã là format mới chưa (Dòng 5 cột 6 có chữ "Vật liệu" hoặc tương tự)
                var cellF5 = ws.Cells[5, 6];
                string valF5 = cellF5.Value2?.ToString() ?? "";
                if (valF5.ToLower().Contains("vật liệu") || valF5.ToLower() == "vl")
                {
                    // Đã là format mới, bỏ qua
                    return;
                }

                // Kiểm tra xem đang ở format cũ hay mới
                // Format cũ: Dòng 4 có "STT" ở A4, dòng 5 bắt đầu dữ liệu
                // Format mới: Dòng 4 có "STT" ở A4, dòng 5 có sub-header, dòng 6 bắt đầu dữ liệu
                var cellA5 = ws.Cells[5, 1];
                string valA5 = cellA5.Value2?.ToString() ?? "";
                
                // Nếu A5 có giá trị số (STT dữ liệu) → format cũ, cần shift xuống
                bool isOldFormat = !string.IsNullOrEmpty(valA5) && int.TryParse(valA5, out _);
                
                if (isOldFormat)
                {
                    // Chèn 1 dòng mới tại dòng 5 để shift dữ liệu xuống
                    Range row5 = ws.Rows[5];
                    row5.Insert(XlInsertShiftDirection.xlShiftDown);
                    
                    // Cập nhật STT cho tất cả DongDuToan (tăng thêm 1)
                    foreach (var hm in duToan.DanhSachHangMuc)
                    {
                        foreach (var ct in hm.DanhSachCongTac)
                        {
                            if (ct.STT >= 5) ct.STT += 1;
                        }
                    }
                }
                
                // Ghi lại header format mới
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
                
                // Merge cells
                var app = ws.Application;
                bool oldAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
                try 
                {
                    try { ws.Range["A4:A5"].UnMerge(); } catch { }
                    try { ws.Range["B4:B5"].UnMerge(); } catch { }
                    try { ws.Range["C4:C5"].UnMerge(); } catch { }
                    try { ws.Range["D4:D5"].UnMerge(); } catch { }
                    try { ws.Range["E4:E5"].UnMerge(); } catch { }
                    try { ws.Range["F4:H4"].UnMerge(); } catch { }
                    try { ws.Range["I4:I5"].UnMerge(); } catch { }
                    
                    ws.Range["A4:A5"].Merge();
                    ws.Range["B4:B5"].Merge();
                    ws.Range["C4:C5"].Merge();
                    ws.Range["D4:D5"].Merge();
                    ws.Range["E4:E5"].Merge();
                    ws.Range["F4:H4"].Merge();
                    ws.Range["I4:I5"].Merge();
                }
                finally
                {
                    app.DisplayAlerts = oldAlerts;
                }
                
                Range headerRange = ws.Range[ws.Cells[4, 1], ws.Cells[5, 9]];
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                // Freeze panes at row 5
                ws.Activate();
                app = ws.Application;
                app.ActiveWindow.FreezePanes = false;
                app.ActiveWindow.SplitRow = 5;
                app.ActiveWindow.SplitColumn = 0;
                app.ActiveWindow.FreezePanes = true;
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

        private void XuatBangTongHopDuToan(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_DuToan");
            SetupHeader(ws, "BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG", 5);

            ws.Cells[3, 1] = "TT";
            ws.Cells[3, 2] = "Khoản mục chi phí";
            ws.Cells[3, 3] = "Ký hiệu";
            ws.Cells[3, 4] = "Cách tính";
            ws.Cells[3, 5] = "Giá trị (đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 5]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            int r = 4;
            var kq = duToan.ChiPhiXD;
            if (kq == null) return;

            void AddRow(string tt, string khoiMuc, string kyHieu, string cachTinh, object giaTriOrFormula, bool isBold)
            {
                ws.Cells[r, 1] = tt;
                ws.Cells[r, 2] = khoiMuc;
                ws.Cells[r, 3] = kyHieu;
                ws.Cells[r, 4] = cachTinh;
                
                if (giaTriOrFormula is string strVal && strVal.StartsWith("=")) {
                    ws.Cells[r, 5].Formula = strVal;
                } else {
                    ws.Cells[r, 5] = giaTriOrFormula;
                }
                
                if (isBold) ws.Range[ws.Cells[r, 1], ws.Cells[r, 5]].Font.Bold = true;
                r++;
            }

            string linkVL = duToan.BangTongHop.DanhSachVatLieu.Count > 0 ? $"='TH_VatLieu'!K{4 + duToan.BangTongHop.DanhSachVatLieu.Count}" : "0";
            string linkNC = duToan.BangTongHop.DanhSachNhanCong.Count > 0 ? $"='TH_NhanCong'!G{4 + duToan.BangTongHop.DanhSachNhanCong.Count}" : "0";
            string linkM = duToan.BangTongHop.DanhSachMay.Count > 0 ? $"='TH_CaMay'!G{4 + duToan.BangTongHop.DanhSachMay.Count}" : "0";

            AddRow("I", "Chi phí trực tiếp", "T", "VL + NC + M", "=SUM(E5:E7)", true);
            AddRow("1", "Chi phí vật liệu", "VL", "Σ(KL × ĐG_VL)", linkVL, false);
            AddRow("2", "Chi phí nhân công", "NC", "Σ(KL × ĐG_NC)", linkNC, false);
            AddRow("3", "Chi phí máy", "M", "Σ(KL × ĐG_M)", linkM, false);
            
            AddRow("II", "Chi phí gián tiếp", "GT", "CPC + TT", "=E9+E10", true);
            AddRow("1", "Chi phí chung", "CPC", $"T × {kq.TiLeCPC}%", $"=E4*{kq.TiLeCPC.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100", false);
            AddRow("2", "Chi phí không xác định được KL từ TK", "TT", $"T × {kq.TiLeTT}%", $"=E4*{kq.TiLeTT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100", false);
            
            AddRow("III", "Thu nhập chịu thuế tính trước", "TL", $"(T + GT) × {kq.TiLeTNCTTT}%", $"=(E4+E8)*{kq.TiLeTNCTTT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100", true);
            
            AddRow("IV", "Chi phí xây dựng trước thuế", "G", "T + GT + TL", "=E4+E8+E11", true);
            
            AddRow("V", "Thuế giá trị gia tăng", "GTGT", $"G × {kq.TiLeGTGT}%", $"=E12*{kq.TiLeGTGT.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100", true);
            
            AddRow("VI", "Chi phí xây dựng sau thuế", "Gxd", "G + GTGT", "=E12+E13", true);
            
            AddRow("VII", "Chi phí nhà tạm", "LT", $"Gxd × {kq.TiLeNhaTam}%", $"=E14*{kq.TiLeNhaTam.ToString(System.Globalization.CultureInfo.InvariantCulture)}/100", true);
            
            AddRow("VIII", "TỔNG CHI PHÍ XÂY DỰNG", "GXD", "Gxd + LT", "=E14+E15", true);

            DrawTableBorders(ws, 3, 1, r - 1, 5);
            ws.Range["E:E"].NumberFormat = "#,##0";
            ws.Columns.AutoFit();
            
            // Di chuyển TH_DuToan lên vị trí đầu tiên (trước sheet tiên lượng)
            ws.Move(Before: wb.Sheets[1]);
        }

        private void XuatPhanTichDonGia(Workbook wb, DuToan duToan, Worksheet wsDuToan)
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

            var dictVL = duToan.BangTongHop.DanhSachVatLieu.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);
            var dictNC = duToan.BangTongHop.DanhSachNhanCong.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);
            var dictMay = duToan.BangTongHop.DanhSachMay.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);

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
                                    ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_VatLieu'!$B:$J, 9, 0), 0)";
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
                            if (wsDuToan != null && ct.STT > 0)
                            {
                                wsDuToan.Cells[ct.STT, 6].Formula = $"='PhanTich_DonGia'!G{r}";
                            }
                            subTotalRows.Add(r);
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
                                ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_NhanCong'!$B:$F, 5, 0), 0)";
                                ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                                r++;
                            }
                            ws.Cells[r, 3] = "Cộng chi phí Nhân công";
                            ws.Cells[r, 7].Formula = $"=SUM(G{r - ncHaoPhi.Count}:G{r - 1})";
                            ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                            if (wsDuToan != null && ct.STT > 0)
                            {
                                wsDuToan.Cells[ct.STT, 7].Formula = $"='PhanTich_DonGia'!G{r}";
                            }
                            subTotalRows.Add(r);
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
                                    ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_CaMay'!$B:$F, 5, 0), 0)";
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
                            if (wsDuToan != null && ct.STT > 0)
                            {
                                wsDuToan.Cells[ct.STT, 8].Formula = $"='PhanTich_DonGia'!G{r}";
                            }
                            subTotalRows.Add(r);
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
                }
            }

            DrawTableBorders(ws, 3, 1, r - 1, 7);
            ws.Range["F:F"].NumberFormat = "#,##0";
            ws.Range["G:G"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
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

                // Nếu không có cấu hình và không nhận diện được và cước = 0, bỏ qua
                if (dm == null && vl.CuocVCOTo <= 0 && (cfg == null || cfg.TongCuLyKm <= 0))
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
                ws.Cells[r, 9].Formula = $"=IFERROR(VLOOKUP(F{r}, 'TH_CaMay'!$B:$F, 5, 0), {donGiaCaMay.ToString(System.Globalization.CultureInfo.InvariantCulture)})";
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
                foreach (var vl in materialsWithBo)
                {
                    var cfgBo = VanChuyenStorage.GetConfigBo(vl.TenVatTu);
                    var dmBo = DinhMucVanChuyenDatabase.NhanDienBo(vl.TenVatTu) ?? DinhMucVanChuyenDatabase.DanhSachBo.First();

                    decimal cuLyMet = cfgBo != null && cfgBo.CuLyMet > 0 ? cfgBo.CuLyMet : 30m;
                    decimal heSoDiaHinh = cfgBo != null && cfgBo.HeSoDiaHinh > 0 ? cfgBo.HeSoDiaHinh : 1.0m;
                    int soTang = cfgBo != null && cfgBo.SoTang > 0 ? cfgBo.SoTang : 1;

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
                    ws.Cells[r, 3].Formula = $"=IFERROR(VLOOKUP(\"N001\", 'TH_NhanCong'!$B:$F, 5, 0), 285000)";
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

        private void XuatBangTongHopVatLieu(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_VatLieu");
            SetupHeader(ws, "BẢNG TỔNG HỢP VẬT LIỆU", 11);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã vật liệu";
            ws.Cells[3, 3] = "Tên vật liệu";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Khối lượng";
            ws.Cells[3, 6] = "Giá mua tại nguồn (đồng)";
            ws.Cells[3, 7] = "Chi phí bốc xếp (đồng)";
            ws.Cells[3, 8] = "Cước VC ô tô (đồng)";
            ws.Cells[3, 9] = "Cước VC bộ (đồng)";
            ws.Cells[3, 10] = "Giá hiện trường (đồng)";
            ws.Cells[3, 11] = "Thành tiền (đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 11]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            int r = 4;
            int stt = 1;
            foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = vl.MaVatTu;
                ws.Cells[r, 3] = vl.TenVatTu;
                ws.Cells[r, 4] = vl.DonVi;
                ws.Cells[r, 5] = (double)vl.TongKhoiLuong;
                ws.Cells[r, 6] = (double)vl.GiaGoc;
                ws.Cells[r, 7] = (double)vl.ChiPhiBocXep;
                ws.Cells[r, 8].Formula = $"=IFERROR(VLOOKUP(B{r}, 'ChietTinh_CuocVC'!$B:$J, 6, 0), {((double)vl.CuocVCOTo).ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                ws.Cells[r, 9].Formula = $"=IFERROR(VLOOKUP(B{r}, 'ChietTinh_CuocVC'!$B:$J, 8, 0), {((double)vl.CuocVCBo).ToString(System.Globalization.CultureInfo.InvariantCulture)})";
                ws.Cells[r, 10].Formula = $"=F{r}+G{r}+H{r}+I{r}";
                ws.Cells[r, 11].Formula = $"=E{r}*J{r}";
                r++;
            }
            
            if (r > 4)
            {
                ws.Cells[r, 3] = "TỔNG CỘNG";
                ws.Cells[r, 11].Formula = $"=SUM(K4:K{r - 1})";
                ws.Range[ws.Cells[r, 3], ws.Cells[r, 11]].Font.Bold = true;
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 11);
            ws.Range["E:E"].NumberFormat = "#,##0.000";
            ws.Range["F:K"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
        }

        private void XuatBangTongHopNhanCong(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_NhanCong");
            SetupHeader(ws, "BẢNG TỔNG HỢP NHÂN CÔNG", 7);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã nhân công";
            ws.Cells[3, 3] = "Tên nhân công";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Khối lượng";
            ws.Cells[3, 6] = "Giá hiện trường (đồng)";
            ws.Cells[3, 7] = "Thành tiền (đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 7]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            int r = 4;
            int stt = 1;
            foreach (var nc in duToan.BangTongHop.DanhSachNhanCong)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = nc.MaVatTu;
                ws.Cells[r, 3] = nc.TenVatTu;
                ws.Cells[r, 4] = nc.DonVi;
                ws.Cells[r, 5] = (double)nc.TongKhoiLuong;
                ws.Cells[r, 6] = (double)nc.GiaHienTruong;
                ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                r++;
            }
            
            if (r > 4)
            {
                ws.Cells[r, 3] = "TỔNG CỘNG";
                ws.Cells[r, 7].Formula = $"=SUM(G4:G{r - 1})";
                ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 7);
            ws.Range["E:E"].NumberFormat = "#,##0.000";
            ws.Range["F:G"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
        }

        private void XuatBangTongHopCaMay(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_CaMay");
            SetupHeader(ws, "BẢNG TỔNG HỢP MÁY THI CÔNG", 7);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã ca máy";
            ws.Cells[3, 3] = "Tên loại máy";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Khối lượng";
            ws.Cells[3, 6] = "Giá hiện trường (đồng)";
            ws.Cells[3, 7] = "Thành tiền (đồng)";

            var headerRange = ws.Range[ws.Cells[3, 1], ws.Cells[3, 7]];
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            int r = 4;
            int stt = 1;
            foreach (var m in duToan.BangTongHop.DanhSachMay)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = m.MaVatTu;
                ws.Cells[r, 3] = m.TenVatTu;
                ws.Cells[r, 4] = m.DonVi;
                ws.Cells[r, 5] = (double)m.TongKhoiLuong;
                ws.Cells[r, 6] = (double)m.GiaHienTruong;
                ws.Cells[r, 7].Formula = $"=E{r}*F{r}";
                r++;
            }
            
            if (r > 4)
            {
                ws.Cells[r, 3] = "TỔNG CỘNG";
                ws.Cells[r, 7].Formula = $"=SUM(G4:G{r - 1})";
                ws.Range[ws.Cells[r, 3], ws.Cells[r, 7]].Font.Bold = true;
                r++;
            }

            DrawTableBorders(ws, 3, 1, r - 1, 7);
            ws.Range["E:E"].NumberFormat = "#,##0.000";
            ws.Range["F:G"].NumberFormat = "#,##0";
            ws.Columns[4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ws.Columns.AutoFit();
        }

        private void XuatBangHeSoDieuChinh(Workbook wb, DuToan duToan)
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
    }
}
