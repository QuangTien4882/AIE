using System;
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

                // Bảng 3, 4, 5: Hao phí (Tạo trước để các bảng sau link công thức không báo lỗi)
                XuatBangTongHopVatLieu(wb, duToan);
                XuatBangTongHopNhanCong(wb, duToan);
                XuatBangTongHopCaMay(wb, duToan);

                // Cập nhật header DuToan sang format 2 dòng nếu cần
                if (wsDuToan != null)
                {
                    ReformatDuToanHeader(wsDuToan, duToan);
                }

                // Bảng 1: Bảng tổng hợp dự toán (TH_DuToan)
                XuatBangTongHopDuToan(wb, duToan);

                // Bảng 2: Bảng phân tích đơn giá chi tiết (PhanTich_DonGia)
                XuatPhanTichDonGia(wb, duToan, wsDuToan);

                // Sắp xếp lại thứ tự sheet: PhanTich_DonGia nằm ngay bên phải DuToan
                if (wsDuToan != null)
                {
                    try
                    {
                        var wsPhanTich = wb.Sheets["PhanTich_DonGia"] as Worksheet;
                        if (wsPhanTich != null)
                        {
                            wsPhanTich.Move(After: wsDuToan);
                        }
                    }
                    catch { }
                }

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

            string linkVL = duToan.BangTongHop.DanhSachVatLieu.Count > 0 ? $"='TH_VatLieu'!G{4 + duToan.BangTongHop.DanhSachVatLieu.Count}" : "0";
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
                                    ws.Cells[r, 6].Formula = $"=IFERROR(VLOOKUP(B{r}, 'TH_VatLieu'!$B:$F, 5, 0), 0)";
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

        private void XuatBangTongHopVatLieu(Workbook wb, DuToan duToan)
        {
            var ws = CreateOrGetSheet(wb, "TH_VatLieu");
            SetupHeader(ws, "BẢNG TỔNG HỢP VẬT LIỆU", 7);

            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã vật liệu";
            ws.Cells[3, 3] = "Tên vật liệu";
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
            foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
            {
                ws.Cells[r, 1] = stt++;
                ws.Cells[r, 2] = vl.MaVatTu;
                ws.Cells[r, 3] = vl.TenVatTu;
                ws.Cells[r, 4] = vl.DonVi;
                ws.Cells[r, 5] = (double)vl.TongKhoiLuong;
                ws.Cells[r, 6] = (double)vl.GiaHienTruong;
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
