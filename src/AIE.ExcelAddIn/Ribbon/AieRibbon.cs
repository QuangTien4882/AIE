using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using AIE.Core.Models;
using AIE.Data;
using AIE.Data.ImportExport;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Forms;
using Dapper;

namespace AIE.ExcelAddIn.Ribbon
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }
        public IntPtr Handle { get; }
    }

    [ComVisible(true)]
    public class AieRibbon : ExcelRibbon
    {
        /// <summary>Dự toán đang làm việc (in-memory). Dùng chung giữa các form.</summary>
        public static AIE.Core.Models.DuToan CurrentDuToan { get; set; }
        /// <summary>Đường dẫn file .dt hiện tại (null nếu chưa lưu)</summary>
        public static string CurrentFilePath { get; set; }
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabAIE' label='AIE Dự Toán'>
                    
                    <group id='groupThietLap' label='Thiết lập dữ liệu'>
                      <button id='btnImport' label='Nhập Database' screentip='Nhập Database' size='normal' showImage='false' onAction='OnImportClicked' />
                      <button id='btnTraCuu' label='Tra cứu định mức' screentip='Tra cứu định mức' size='normal' showImage='false' onAction='OnTraCuuClicked' />
                      <button id='btnTaoTemplate' label='Tạo File mẫu' screentip='Tạo File mẫu' size='normal' showImage='false' onAction='OnTaoTemplateClicked' />
                      <button id='btnXuatDonGia' label='Trích xuất Đơn Giá' screentip='Trích xuất Đơn Giá' size='normal' showImage='false' onAction='OnXuatDonGiaClicked' />
                      <button id='btnQuanLyDonGia' label='Quản lý Đơn giá' screentip='Quản lý Đơn giá' size='normal' showImage='false' onAction='OnQuanLyDonGiaClicked' />
                      <button id='btnDeleteDb' label='Xóa Database' screentip='Xóa toàn bộ Database' size='normal' showImage='false' onAction='OnDeleteDbClicked' />
                    </group>

                    <group id='groupThamDinh' label='Thẩm định dự toán'>
                      <button id='btnDonGiaThamDinh' label='Đơn giá thẩm định' screentip='Lập Bộ Đơn giá thẩm định' size='normal' showImage='false' onAction='OnDonGiaThamDinhClicked' />
                      <button id='btnMoDonGiaThamDinh' label='Mở Bộ Đơn giá' screentip='Mở lại Bộ Đơn giá Thẩm định' size='normal' showImage='false' onAction='OnMoDonGiaThamDinhClicked' />
                      <button id='btnKiemTra' label='Kiểm tra' screentip='Kiểm tra dự toán' size='normal' showImage='false' onAction='OnKiemTraClicked' />
                      <button id='btnBaoCaoTD' label='Xuất Báo cáo' screentip='Xuất Báo cáo' size='normal' showImage='false' onAction='OnBaoCaoTDClicked' />
                    </group>

                    <group id='groupLapDuToan' label='Lập dự toán'>
                      <button id='btnTaoDuToanMoi' label='Tạo Dự toán mới' screentip='Tạo Dự toán mới' size='normal' showImage='false' onAction='OnTaoDuToanMoiClicked' />
                      <button id='btnGoiDonGia' label='Gọi Đơn giá' screentip='Gọi Đơn giá' size='normal' showImage='false' onAction='OnGoiDonGiaClicked' />
                      <button id='btnTinhGiaHienTruong' label='Giá VL, NC, MTC' screentip='Giá VL, NC, MTC' size='normal' showImage='false' onAction='OnTinhGiaHienTruongClicked' />
                      <button id='btnTinhTongHop' label='Tính Tổng hợp' screentip='Tính Tổng hợp' size='normal' showImage='false' onAction='OnTinhTongHopClicked' />
                    </group>

                    <group id='groupFile' label='File Dự toán'>
                      <button id='btnLuuDuToan' label='Lưu Dự toán' screentip='Lưu Dự toán (.dt)' size='normal' showImage='false' onAction='OnLuuDuToanClicked' />
                      <button id='btnMoDuToan' label='Mở Dự toán' screentip='Mở Dự toán (.dt)' size='normal' showImage='false' onAction='OnMoDuToanClicked' />
                    </group>

                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnImportClicked(IRibbonControl control)
        {
            try
            {
                var form = new ImportDatabaseForm();
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở form Import: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnXuatDonGiaClicked(IRibbonControl control)
        {
            try
            {
                using var fbd = new SaveFileDialog();
                fbd.Filter = "Excel Files|*.xlsx";
                fbd.Title = "Lưu file Đơn giá chuẩn";
                fbd.FileName = "DonGiaChuan_Template.xlsx";
                
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    var db = new DatabaseManager();
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var haoPhis = repo.GetAllHaoPhi().ToList();

                    if (haoPhis.Count == 0)
                    {
                        MessageBox.Show("Không có dữ liệu Hao phí trong hệ thống. Vui lòng nạp Định mức trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var generator = new ExcelTemplateGenerator();
                    generator.ExportDonGiaTemplate(fbd.FileName, haoPhis);
                    
                    MessageBox.Show("Đã trích xuất danh mục Đơn giá thành công!\nBạn hãy mở file lên, điền giá trị và dùng nó cho các lần thẩm định sau.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất Đơn giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTraCuuClicked(IRibbonControl control)
        {
            try
            {
                var form = new TraCuuDinhMucForm();
                // Show floating over Excel without blocking
                form.Show(new WindowWrapper(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở form Tra cứu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTaoTemplateClicked(IRibbonControl control)
        {
            try
            {
                var dialog = new FolderBrowserDialog();
                dialog.Description = "Chọn thư mục để lưu các file mẫu (Templates):";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var generator = new ExcelTemplateGenerator();
                    generator.GenerateAllTemplates(dialog.SelectedPath);
                    MessageBox.Show(
                        "Đã tạo 6 file mẫu thành công tại:\n" + dialog.SelectedPath,
                        "AIE Dự Toán",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi Tạo File mẫu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDeleteDbClicked(IRibbonControl control)
        {
            var form = new Forms.DeleteDatabaseForm();
            form.ShowDialog();
        }

        public void OnKiemTraClicked(IRibbonControl control)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không có file Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sheetNames = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                var form = new ThamDinhConfigForm(sheetNames);
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var config = form.ResultConfig;
                    
                    // 1. Đọc dữ liệu
                    var reader = new AIE.ExcelAddIn.Services.DuToanExcelReader();
                    var danhSachCongTac = reader.Read(config);

                    if (danhSachCongTac.Count == 0)
                    {
                        MessageBox.Show("Không tìm thấy dữ liệu công tác nào. Vui lòng Kiểm tra lại cấu hình cột.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 2. Thẩm định
                    var db = new DatabaseManager(); // Đảm bảo đã init db
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                    var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                    var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                    var engine = new AIE.ExcelAddIn.Services.ThamDinhEngine(repo, vlRepo, ncRepo, mayRepo);
                    
                    var ketQua = engine.KiemTra(danhSachCongTac, form.SelectedBoDonGiaId);

                    // 3. Xuất kết quả (gộp Định mức + Đơn giá + Thành tiền vào 1 sheet)
                    var writer = new AIE.ExcelAddIn.Services.ThamDinhExcelWriter();
                    writer.ExportResult(config, ketQua);
                    
                    int soLoi = ketQua.Sum(x => x.DanhSachSaiLech.Count(s => !string.IsNullOrEmpty(s.LoaiLoi)));
                    MessageBox.Show($"Thẩm định hoàn tất!\nĐã Kiểm tra: {ketQua.Count} công tác.\nPhát hiện: {soLoi} sai lệch.\nKết quả đã được xuất ra sheet KQ_ThamDinh.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi Kiểm tra thẩm định: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDonGiaThamDinhClicked(IRibbonControl control)
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không có file Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sheetNames = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                // Tạm thời mượn ThamDinhConfigForm để người dùng map cột
                var form = new ThamDinhConfigForm(sheetNames);
                form.Text = "Cấu Hình Đọc Dự Toán - Lấy Đơn Giá";
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var config = form.ResultConfig;
                    
                    // 1. Đọc dữ liệu
                    var reader = new AIE.ExcelAddIn.Services.DuToanExcelReader();
                    var danhSachCongTac = reader.Read(config);

                    if (danhSachCongTac.Count == 0)
                    {
                        MessageBox.Show("Không tìm thấy dữ liệu công tác nào. Vui lòng kiểm tra lại cấu hình cột.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 2. Thẩm định để map với định mức chuẩn
                    var db = new DatabaseManager();
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                    var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                    var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                    var engine = new AIE.ExcelAddIn.Services.ThamDinhEngine(repo, vlRepo, ncRepo, mayRepo);
                    
                    var ketQua = engine.KiemTra(danhSachCongTac);

                    // 3. Trích xuất vật tư
                    var danhSachVatTu = engine.TrichXuatVatTu(ketQua);
                    
                    // 4. Mở Form ThamDinhDonGiaForm
                    var loading = new AIE.ExcelAddIn.Forms.LoadingForm("Đang mở Thẩm định đơn giá...");
                    loading.Show();
                    System.Windows.Forms.Application.DoEvents();
                    
                    var donGiaForm = new ThamDinhDonGiaForm(danhSachVatTu, config.Vung);
                    
                    loading.Close();
                    loading.Dispose();
                    
                    if (donGiaForm.ShowDialog() == DialogResult.OK && donGiaForm.SavedBoDonGiaId.HasValue)
                    {
                        var result = MessageBox.Show("Bạn có muốn áp dụng Bộ đơn giá vừa tạo để Thẩm định (Kiểm tra) dự toán này ngay không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            var loadingKiemTra = new AIE.ExcelAddIn.Forms.LoadingForm("Đang thẩm định lại...");
                            loadingKiemTra.Show();
                            System.Windows.Forms.Application.DoEvents();
                            
                            var ketQuaMoi = engine.KiemTra(danhSachCongTac, donGiaForm.SavedBoDonGiaId);
                            var writer = new AIE.ExcelAddIn.Services.ThamDinhExcelWriter();
                            writer.ExportResult(config, ketQuaMoi);
                            
                            loadingKiemTra.Close();
                            loadingKiemTra.Dispose();
                            
                            MessageBox.Show("Đã hoàn tất thẩm định và xuất kết quả ra sheet KQ_ThamDinh.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đọc dự toán: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnMoDonGiaThamDinhClicked(IRibbonControl control)
        {
            try
            {
                var db = new DatabaseManager();
                var repo = new BoDonGiaRepository(db.Context);
                
                using (var form = new ChonBoDonGiaForm(repo))
                {
                    form.Text = "Mở Bộ Đơn Giá Thẩm Định";
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedBoDonGiaId > 0)
                    {
                        var loading = new LoadingForm("Đang tải dữ liệu bộ đơn giá...");
                        loading.Show();
                        System.Windows.Forms.Application.DoEvents();

                        int boId = form.SelectedBoDonGiaId;
                        
                        // Lấy vùng của bộ đơn giá
                        var selectedBo = repo.GetAll().FirstOrDefault(x => x.Id == boId);
                        AIE.Core.Enums.Vung vung = AIE.Core.Enums.Vung.VungII; // Mặc định vì BoDonGia hiện không lưu Vùng

                        var vlRepo = new VatLieuRepository(db.Context);
                        var ncRepo = new NhanCongRepository(db.Context);
                        var mayRepo = new MayThiCongRepository(db.Context);

                        // Reconstruct VatTuGiaModel list from the saved sets
                        var dsVatTu = new System.Collections.Generic.List<VatTuGiaModel>();
                        
                        var giaVL = repo.GetGiaVL(boId);
                        using (var conn = db.Context.GetConnection())
                        {
                            foreach(var vl in giaVL)
                            {
                                var master = vlRepo.GetByMa(vl.MaVL);
                                string name = master?.TenVL ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = vl.MaVL }) ?? vl.MaVL;
                                dsVatTu.Add(new VatTuGiaModel { MaHieu = vl.MaVL, TenVatTu = name, DonVi = master?.DonVi ?? "", GiaChuan = vl.GiaGoc, LoaiHP = AIE.Core.Enums.LoaiHaoPhi.VL });
                            }

                            var giaNC = repo.GetGiaNC(boId);
                            foreach(var nc in giaNC)
                            {
                                var master = ncRepo.GetByMa(nc.MaNC);
                                string name = master?.TenNC ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = nc.MaNC }) ?? nc.MaNC;
                                dsVatTu.Add(new VatTuGiaModel { MaHieu = nc.MaNC, TenVatTu = name, DonVi = master?.DonVi ?? "", GiaChuan = nc.DonGia, LoaiHP = AIE.Core.Enums.LoaiHaoPhi.NC });
                            }

                            var giaMay = repo.GetGiaMay(boId);
                            foreach(var m in giaMay)
                            {
                                var master = mayRepo.GetByMa(m.MaMay);
                                string name = master?.TenMay ?? conn.QueryFirstOrDefault<string>("SELECT TenHaoPhi FROM HaoPhi WHERE MaHieuHP = @Ma", new { Ma = m.MaMay }) ?? m.MaMay;
                                dsVatTu.Add(new VatTuGiaModel { MaHieu = m.MaMay, TenVatTu = name, DonVi = master?.DonVi ?? "", GiaChuan = m.DonGia, LoaiHP = AIE.Core.Enums.LoaiHaoPhi.MAY });
                            }
                        }

                        var donGiaForm = new ThamDinhDonGiaForm(dsVatTu, vung, boId);
                        if (selectedBo != null)
                        {
                            donGiaForm.SetTenBoDonGia(selectedBo.TenBo, selectedBo.GiaXang, selectedBo.GiaDiezel, selectedBo.GiaDien);
                        }

                        loading.Close();
                        loading.Dispose();

                        // Không tự động chạy Kiểm tra khi mở lại, người dùng lưu xong tự ấn Kiểm tra
                        donGiaForm.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở Bộ Đơn Giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnQuanLyDonGiaClicked(IRibbonControl control)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var db = new AIE.Data.DatabaseManager();
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var mayDmRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);

                var loading = new AIE.ExcelAddIn.Forms.LoadingForm("Đang mở Quản lý đơn giá...");
                loading.Show();
                System.Windows.Forms.Application.DoEvents();

                var form = new AIE.ExcelAddIn.Forms.QuanLyDonGiaForm(vlRepo, ncRepo, mayRepo, mayDmRepo);
                Cursor.Current = Cursors.Default;
                
                loading.Close();
                loading.Dispose();
                
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnBaoCaoTDClicked(IRibbonControl control)
        {
            MessageBox.Show("Chức năng Xuất Báo cáo thẩm định đang được phát triển.", "AIE Dự Toán");
        }

        public void OnTaoDuToanMoiClicked(IRibbonControl control)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                
                // Check if there's an active workbook that is not completely empty
                if (wb != null)
                {
                    var ws = wb.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                    bool isEmpty = false;
                    if (ws != null && ws.Name.StartsWith("Sheet"))
                    {
                        var range = ws.UsedRange;
                        var cell1 = ws.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                        string cellText = cell1 != null ? (cell1.Text?.ToString() ?? "") : "";
                        if (range.Rows.Count <= 1 && range.Columns.Count <= 1 && string.IsNullOrEmpty(cellText))
                        {
                            isEmpty = true;
                        }
                    }

                    if (!isEmpty)
                    {
                        var result = MessageBox.Show(
                            "Đang có dự án mở. Bạn có muốn lưu dự án hiện hành trước khi tạo mới không?\n\n- Chọn Yes để Lưu dự án cũ và Tạo mới\n- Chọn No để Đóng dự án cũ (không lưu) và Tạo mới\n- Chọn Cancel để Hủy thao tác",
                            "Xác nhận đóng dự án",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Cancel)
                        {
                            return; // Hủy tạo mới
                        }
                        else if (result == DialogResult.Yes)
                        {
                            bool saved = SaveDuToan();
                            if (!saved) return; // Nếu người dùng hủy lưu thì không đóng
                            
                            wb.Close(false);
                        }
                        else if (result == DialogResult.No)
                        {
                            wb.Close(false); // Đóng không lưu
                        }
                    }
                }

                var form = new TaoDuToanMoiForm();
                form.ShowDialog(new WindowWrapper(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khởi tạo dự toán: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnGoiDonGiaClicked(IRibbonControl control)
        {
            try
            {
                var app = (Application)ExcelDnaUtil.Application;
                var ws = app.ActiveSheet as Worksheet;
                if (ws == null) return;

                var selection = app.Selection as Range;
                if (selection == null) return;

                // Create DB context and repo
                var dbManager = new DatabaseManager();
                var repo = new AIE.Data.Repositories.CongTacRepository(dbManager.Context);

                int updatedCount = 0;
                
                // Get all rows in the selection
                foreach (Range row in selection.Rows)
                {
                    int rowIndex = row.Row;
                    
                    // Assume Mã hiệu is in Column 2 (B) based on our template
                    var maHieuCell = ws.Cells[rowIndex, 2] as Range;
                    if (maHieuCell != null && maHieuCell.Value2 != null)
                    {
                        string maHieu = maHieuCell.Value2.ToString().Trim();
                        if (!string.IsNullOrEmpty(maHieu))
                        {
                            var congTac = repo.GetByMaHieu(maHieu);
                            if (congTac != null)
                            {
                                ws.Cells[rowIndex, 3].Value2 = congTac.TenCongTac;
                                ws.Cells[rowIndex, 4].Value2 = congTac.DonVi;
                                
                                // Calculate STT if empty
                                var sttCell = ws.Cells[rowIndex, 1].Value2;
                                if (sttCell == null || string.IsNullOrWhiteSpace(sttCell.ToString()))
                                {
                                    int stt = 1;
                                    for (int i = rowIndex - 1; i >= 5; i--) // row 4 is header
                                    {
                                        var prevCell = ws.Cells[i, 1].Value2;
                                        int prevStt = 0;
                                        if (prevCell != null && int.TryParse(prevCell.ToString(), out prevStt))
                                        {
                                            stt = prevStt + 1;
                                            break;
                                        }
                                    }
                                    ws.Cells[rowIndex, 1].Value2 = stt;
                                }

                                // Apply borders and alignment
                                var rowRange = ws.Range[ws.Cells[rowIndex, 1], ws.Cells[rowIndex, 9]];
                                rowRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                rowRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                                updatedCount++;
                            }
                        }
                    }
                }

                if (updatedCount == 0)
                {
                    MessageBox.Show("Không tìm thấy Mã hiệu nào hợp lệ trong vùng đang chọn. Vui lòng chọn các ô ở cột Mã hiệu (cột B).", "AIE Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi Gọi Đơn giá: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTinhTongHopClicked(IRibbonControl control)
        {
            try
            {
                var db = new DatabaseManager();
                var ctRepo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var dinhMucMayRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var duToan = excelService.ReadBOQFromActiveSheet();

                if (duToan.DanhSachHangMuc[0].DanhSachCongTac.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy công tác nào trong file Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Chạy phân tích vật tư để có danh sách vật liệu, nhân công, máy
                var service = new AIE.ExcelAddIn.Services.PhanTichVatTuService(ctRepo, mayRepo, dinhMucMayRepo, vlRepo, ncRepo);
                service.PhanTich(duToan);

                // Tính toán đơn giá chi tiết (Sử dụng giá gốc trong DB hoặc giá đã áp ở bước trước)
                var phanTichDonGiaService = new AIE.ExcelAddIn.Services.PhanTichDonGiaService(ctRepo);
                phanTichDonGiaService.TinhDonGiaChiTiet(duToan);
                
                // Mở Form Tính Tổng hợp
                using var form = new TinhTongHopForm(duToan, excelService);
                form.ShowDialog();
                
                CurrentDuToan = duToan;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTinhGiaHienTruongClicked(IRibbonControl control)
        {
            try
            {
                var db = new DatabaseManager();
                var ctRepo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                var dinhMucMayRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
                var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var duToan = excelService.ReadBOQFromActiveSheet();

                if (duToan.DanhSachHangMuc[0].DanhSachCongTac.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy công tác nào trong file Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Giữ lại BoDonGiaId nếu đã mở từ file .dt trước đó
                if (CurrentDuToan != null && CurrentDuToan.BoDonGiaId.HasValue)
                {
                    duToan.BoDonGiaId = CurrentDuToan.BoDonGiaId;
                }
                else
                {
                    var boDonGiaRepo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                    using var frmBoDonGia = new Forms.ChonBoDonGiaForm(boDonGiaRepo);
                    if (frmBoDonGia.ShowDialog() == DialogResult.OK)
                    {
                        duToan.BoDonGiaId = frmBoDonGia.SelectedBoDonGiaId;
                    }
                    else
                    {
                        return; // Hủy tính giá nếu không chọn bộ đơn giá
                    }
                }

                var service = new AIE.ExcelAddIn.Services.PhanTichVatTuService(ctRepo, mayRepo, dinhMucMayRepo, vlRepo, ncRepo);
                service.PhanTich(duToan);

                // Load giá từ Bộ Đơn Giá vừa chọn (Ghi đè giá gốc từ thư viện chung)
                var boDonGiaRepo2 = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                var giaVl = boDonGiaRepo2.GetGiaVL(duToan.BoDonGiaId.Value).GroupBy(x => x.MaVL).ToDictionary(g => g.Key, g => g.First());
                var giaNc = boDonGiaRepo2.GetGiaNC(duToan.BoDonGiaId.Value).GroupBy(x => x.MaNC).ToDictionary(g => g.Key, g => g.First());
                var giaMay = boDonGiaRepo2.GetGiaMay(duToan.BoDonGiaId.Value).GroupBy(x => x.MaMay).ToDictionary(g => g.Key, g => g.First());

                foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
                    if (giaVl.TryGetValue(vl.MaVatTu, out var g)) { vl.GiaGoc = g.GiaGoc; vl.CuocVanChuyen = g.CuocVC; }
                
                foreach (var nc in duToan.BangTongHop.DanhSachNhanCong)
                    if (giaNc.TryGetValue(nc.MaVatTu, out var g)) nc.GiaGoc = g.DonGia;
                
                foreach (var may in duToan.BangTongHop.DanhSachMay)
                    if (giaMay.TryGetValue(may.MaVatTu, out var g)) may.GiaGoc = g.DonGia;

                var phanTichDonGiaService = new AIE.ExcelAddIn.Services.PhanTichDonGiaService(ctRepo);

                // Lưu vào bộ nhớ chung để có thể Save ra file .dt
                CurrentDuToan = duToan;

                var loading = new AIE.ExcelAddIn.Forms.LoadingForm("Đang mở Giá VL, NC, MTC...");
                loading.Show();
                System.Windows.Forms.Application.DoEvents();
                
                using var form = new TinhGiaHienTruongForm(duToan, service, phanTichDonGiaService);
                
                loading.Close();
                loading.Dispose();
                
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ==================== LƯU / MỞ FILE DỰ TOÁN (.dt) ====================

        public void OnLuuDuToanClicked(IRibbonControl control)
        {
            SaveDuToan();
        }

        public bool SaveDuToan()
        {
            try
            {
                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                var latestDuToan = excelService.ReadBOQFromActiveSheet();
                
                if (CurrentDuToan != null)
                {
                    latestDuToan.BoDonGiaId = CurrentDuToan.BoDonGiaId;
                    latestDuToan.ChiPhiXD = CurrentDuToan.ChiPhiXD;
                    latestDuToan.BangTongHop = CurrentDuToan.BangTongHop;
                    latestDuToan.LoaiCongTrinh = CurrentDuToan.LoaiCongTrinh;
                    
                    // Bảo toàn DanhSachHaoPhi từ CurrentDuToan
                    var oldCongTacDict = new System.Collections.Generic.Dictionary<string, AIE.Core.Models.DongDuToan>();
                    foreach (var hm in CurrentDuToan.DanhSachHangMuc)
                    {
                        foreach (var ct in hm.DanhSachCongTac)
                        {
                            if (!string.IsNullOrEmpty(ct.MaHieu) && !oldCongTacDict.ContainsKey(ct.MaHieu))
                                oldCongTacDict[ct.MaHieu] = ct;
                        }
                    }
                    
                    foreach (var hm in latestDuToan.DanhSachHangMuc)
                    {
                        foreach (var ct in hm.DanhSachCongTac)
                        {
                            if (!string.IsNullOrEmpty(ct.MaHieu) && oldCongTacDict.TryGetValue(ct.MaHieu, out var oldCt))
                            {
                                ct.DanhSachHaoPhi = oldCt.DanhSachHaoPhi;
                            }
                        }
                    }
                }
                CurrentDuToan = latestDuToan;

                string filePath = CurrentFilePath;

                // Nếu chưa có đường dẫn, hỏi người dùng
                if (string.IsNullOrEmpty(filePath))
                {
                    using var sfd = new SaveFileDialog();
                    sfd.Filter = "File Dự Toán (*.dt)|*.dt";
                    sfd.Title = "Lưu Dự Toán";
                    sfd.FileName = string.IsNullOrEmpty(CurrentDuToan.TenCongTrinh)
                        ? "DuToan_Moi.dt"
                        : CurrentDuToan.TenCongTrinh.Replace(" ", "_") + ".dt";

                    if (sfd.ShowDialog() != DialogResult.OK) return false;
                    filePath = sfd.FileName;
                }

                AIE.ExcelAddIn.Services.DuToanFileService.Save(CurrentDuToan, filePath);
                CurrentFilePath = filePath;

                MessageBox.Show(
                    $"Đã lưu dự toán thành công!\n\nFile: {filePath}",
                    "Lưu Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public void OnMoDuToanClicked(IRibbonControl control)
        {
            try
            {
                using var ofd = new OpenFileDialog();
                ofd.Filter = "File Dự Toán (*.dt)|*.dt|Tất cả file (*.*)|*.*";
                ofd.Title = "Mở Dự Toán";

                if (ofd.ShowDialog() != DialogResult.OK) return;

                var duToan = AIE.ExcelAddIn.Services.DuToanFileService.Load(ofd.FileName);
                CurrentDuToan = duToan;
                CurrentFilePath = ofd.FileName;

                var app = (Application)ExcelDnaUtil.Application;
                app.Workbooks.Add(); // Tạo workbook mới

                // Ghi dữ liệu ra Excel (sheet hiện tại) để người dùng có thể xem/chỉnh sửa
                var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                excelService.WriteBOQToActiveSheet(duToan);

                // Nếu dự toán đã được tính toán tổng hợp (đã có ChiPhiXD), tự động xuất lại các bảng biểu
                if (duToan.ChiPhiXD != null)
                {
                    var xuatService = new AIE.ExcelAddIn.Services.XuatBangBieuService();
                    xuatService.Xuat7BangBieu(duToan);
                }

                MessageBox.Show(
                    $"Đã mở dự toán thành công!\n\nCông trình: {duToan.TenCongTrinh}\nLoại: {duToan.LoaiCongTrinh}\nFile: {ofd.FileName}",
                    "Mở Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
