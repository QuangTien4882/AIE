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
using AIE.ExcelAddIn.Helpers;
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

        // Singleton references: tránh mở nhiều instance cùng lúc (Phương án B: Modeless)
        private static TongHopKinhPhiForm _tongHopForm;
        private static TinhGiaHienTruongForm _tinhGiaForm;
        private static ThamDinhDonGiaForm _thamDinhForm;
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
                      <button id='btnTongHopKinhPhi' label='Tổng hợp kinh phí' screentip='Tổng hợp Chi phí xây dựng, Dự toán &amp; Tổng mức đầu tư' size='normal' showImage='false' onAction='OnTongHopKinhPhiClicked' />
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
                // Show floating without owner to allow Alt+Tab and independent minimize
                form.Show();
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
                    
                    if (_thamDinhForm != null && !_thamDinhForm.IsDisposed)
                    {
                        loading.Close();
                        loading.Dispose();
                        _thamDinhForm.Activate();
                        return;
                    }
                    var donGiaForm = new ThamDinhDonGiaForm(danhSachVatTu, config.Vung);
                    _thamDinhForm = donGiaForm;
                    
                    loading.Close();
                    loading.Dispose();
                    
                    // Xử lý logic sau khi form đóng qua FormClosed event (Modeless)
                    var capturedEngine = engine;
                    var capturedConfig = config;
                    var capturedDanhSachCongTac = danhSachCongTac;
                    donGiaForm.FormClosed += (s, ev) =>
                    {
                        _thamDinhForm = null;
                        if (donGiaForm.DialogResult == DialogResult.OK && donGiaForm.SavedBoDonGiaId.HasValue)
                        {
                            var result = MessageBox.Show("Bạn có muốn áp dụng Bộ đơn giá vừa tạo để Thẩm định (Kiểm tra) dự toán này ngay không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.Yes)
                            {
                                var loadingKiemTra = new AIE.ExcelAddIn.Forms.LoadingForm("Đang thẩm định lại...");
                                loadingKiemTra.Show();
                                System.Windows.Forms.Application.DoEvents();
                                
                                var ketQuaMoi = capturedEngine.KiemTra(capturedDanhSachCongTac, donGiaForm.SavedBoDonGiaId);
                                var writer = new AIE.ExcelAddIn.Services.ThamDinhExcelWriter();
                                writer.ExportResult(capturedConfig, ketQuaMoi);
                                
                                loadingKiemTra.Close();
                                loadingKiemTra.Dispose();
                                
                                MessageBox.Show("Đã hoàn tất thẩm định và xuất kết quả ra sheet KQ_ThamDinh.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    };
                    donGiaForm.Show();
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

                // Create DB context and repo
                var dbManager = new DatabaseManager();
                var repo = new AIE.Data.Repositories.CongTacRepository(dbManager.Context);

                // Tìm phạm vi dòng dữ liệu trong sheet (từ dòng 6)
                int usedMax = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1;
                int maxRow = Math.Max(usedMax, 6);

                // Xác định dòng kết thúc thực sự có dữ liệu
                for (int r = 6; r <= Math.Max(usedMax, 200); r++)
                {
                    string a = ws.Cells[r, 1]?.Value2?.ToString()?.Trim() ?? "";
                    string b = ws.Cells[r, 2]?.Value2?.ToString()?.Trim() ?? "";
                    string c = ws.Cells[r, 3]?.Value2?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(a) || !string.IsNullOrEmpty(b) || !string.IsNullOrEmpty(c))
                    {
                        if (r > maxRow) maxRow = r;
                    }
                }

                // Tập hợp các dòng được chọn (nếu có)
                var selectedRowIndices = new HashSet<int>();
                if (selection != null)
                {
                    foreach (Range r in selection.Rows)
                    {
                        if (r.Row >= 6) selectedRowIndices.Add(r.Row);
                    }
                }

                int updatedCount = 0;
                int currentStt = 0;

                // Duyệt qua TOÀN BỘ các dòng từ dòng 6 đến maxRow để định dạng và tra cứu
                for (int rowIndex = 6; rowIndex <= maxRow; rowIndex++)
                {
                    var rowRange = ws.Range[ws.Cells[rowIndex, 1], ws.Cells[rowIndex, 11]];
                    rowRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    rowRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                    var maHieuCell = ws.Cells[rowIndex, 2] as Range;
                    string maHieu = maHieuCell?.Value2?.ToString()?.Trim() ?? "";
                    string sttText = ws.Cells[rowIndex, 1]?.Value2?.ToString()?.Trim() ?? "";
                    string tenText = ws.Cells[rowIndex, 3]?.Value2?.ToString()?.Trim() ?? "";

                    // Kiểm tra xem dòng có phải là dòng Dữ liệu/Hạng mục không
                    if (string.IsNullOrEmpty(maHieu) && string.IsNullOrEmpty(sttText) && string.IsNullOrEmpty(tenText))
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(maHieu))
                    {
                        // Dòng CÔNG TÁC
                        // Nếu dòng nằm trong selection hoặc tên đang trống thì tra cứu DB
                        bool shouldLookup = selectedRowIndices.Count <= 1 || selectedRowIndices.Contains(rowIndex) || string.IsNullOrEmpty(tenText);
                        if (shouldLookup)
                        {
                            var congTac = repo.GetByMaHieu(maHieu);
                            if (congTac != null)
                            {
                                ws.Cells[rowIndex, 3].Value2 = congTac.TenCongTac;
                                ws.Cells[rowIndex, 4].Value2 = congTac.DonVi;
                                updatedCount++;
                            }
                        }

                        // Đánh số STT tăng dần nếu chưa có
                        if (string.IsNullOrWhiteSpace(sttText) || !int.TryParse(sttText, out _))
                        {
                            currentStt++;
                            ws.Cells[rowIndex, 1].Value2 = currentStt;
                        }
                        else if (int.TryParse(sttText, out int s))
                        {
                            currentStt = s;
                        }

                        // Định dạng cho dòng công tác
                        rowRange.Font.Bold = false;
                        rowRange.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;
                        ws.Cells[rowIndex, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[rowIndex, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        ws.Cells[rowIndex, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                        ws.Cells[rowIndex, 3].WrapText = true;

                        ExcelFormatHelper.ApplyQuantityFormat(ws.Cells[rowIndex, 5], 2);
                        ExcelFormatHelper.ApplyIntegerFormat(ws.Range[ws.Cells[rowIndex, 6], ws.Cells[rowIndex, 11]]);
                        ws.Range[ws.Cells[rowIndex, 5], ws.Cells[rowIndex, 11]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    }
                    else
                    {
                        // Dòng HẠNG MỤC hoặc dòng tiêu đề
                        if (!tenText.StartsWith("TỔNG CỘNG", StringComparison.OrdinalIgnoreCase))
                        {
                            rowRange.Font.Bold = true;
                            rowRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 235, 252));
                            ws.Cells[rowIndex, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            ws.Cells[rowIndex, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                        }
                    }
                }

                // Đóng khung toàn bộ bảng từ dòng 4 đến maxRow
                var wholeTable = ws.Range[ws.Cells[4, 1], ws.Cells[maxRow, 11]];
                wholeTable.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                // Đảm bảo mở lại hiển thị cột 10 (J) nếu vô tình bị ẩn
                try
                {
                    ((Range)ws.Columns[10]).Hidden = false;
                    ((Range)ws.Columns[10]).ColumnWidth = 16;
                }
                catch { }

                if (updatedCount == 0 && selectedRowIndices.Count > 1)
                {
                    MessageBox.Show("Đã chuẩn hóa định dạng bảng và các dòng Hạng mục. Không tìm thấy Mã hiệu mới nào cần tra cứu.", "AIE Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi Gọi Đơn giá: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTongHopKinhPhiClicked(IRibbonControl control)
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
                var duToan = CurrentDuToan;

                if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Count == 0 || duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Count) == 0)
                {
                    try
                    {
                        duToan = excelService.ReadBOQFromActiveSheet();
                    }
                    catch { }
                }

                if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Count == 0 || duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Count) == 0)
                {
                    MessageBox.Show("Không tìm thấy công tác nào trong file Excel hoặc dự toán hiện hành.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Chạy phân tích vật tư nếu chưa có bảng tổng hợp vật tư
                if (duToan.BangTongHop == null || duToan.BangTongHop.DanhSachVatLieu.Count == 0)
                {
                    var service = new AIE.ExcelAddIn.Services.PhanTichVatTuService(ctRepo, mayRepo, dinhMucMayRepo, vlRepo, ncRepo);
                    service.PhanTich(duToan);
                }

                // Tính toán đơn giá chi tiết
                var phanTichDonGiaService = new AIE.ExcelAddIn.Services.PhanTichDonGiaService(ctRepo);
                phanTichDonGiaService.TinhDonGiaChiTiet(duToan);
                
                // Mở Form Tổng hợp kinh phí (Modeless: người dùng có thể click vào Excel)
                if (_tongHopForm != null && !_tongHopForm.IsDisposed)
                {
                    _tongHopForm.Activate();
                    return;
                }
                CurrentDuToan = duToan;
                var form = new TongHopKinhPhiForm(duToan);
                _tongHopForm = form;
                form.FormClosed += (s, ev) => { _tongHopForm = null; };
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnTinhTongHopClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnXuatBangTHClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnTongHopDuToanClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        public void OnTongMucDauTuClicked(IRibbonControl control)
        {
            OnTongHopKinhPhiClicked(control);
        }

        private void MoFormTongHopKinhPhi(bool macDinhTongMucDauTu)
        {
            try
            {
                var duToan = CurrentDuToan;
                if (duToan == null)
                {
                    try
                    {
                        var excelService = new AIE.ExcelAddIn.Services.LapDuToanExcelService();
                        duToan = excelService.ReadBOQFromActiveSheet();
                    }
                    catch { }

                    if (duToan == null)
                    {
                        duToan = new DuToan();
                    }
                }

                if (_tongHopForm != null && !_tongHopForm.IsDisposed)
                {
                    _tongHopForm.Activate();
                    return;
                }
                CurrentDuToan = duToan;
                var form = new TongHopKinhPhiForm(duToan, macDinhTongMucDauTu);
                _tongHopForm = form;
                form.FormClosed += (s, ev) => { _tongHopForm = null; };
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở Bảng Tổng hợp kinh phí: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                if (duToan == null || duToan.DanhSachHangMuc == null || duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Count) == 0)
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

                // Bảo toàn thông tin cước và cấu hình đã tính từ phiên trước nếu có
                if (CurrentDuToan?.BangTongHop?.DanhSachVatLieu != null)
                {
                    var oldVlDict = CurrentDuToan.BangTongHop.DanhSachVatLieu
                        .GroupBy(x => x.MaVatTu)
                        .ToDictionary(g => g.Key, g => g.First());

                    foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
                    {
                        if (oldVlDict.TryGetValue(vl.MaVatTu, out var oldVl))
                        {
                            vl.GiaGoc = oldVl.GiaGoc;
                            vl.ChiPhiBocXep = oldVl.ChiPhiBocXep;
                            vl.CuocVCOTo = oldVl.CuocVCOTo;
                            vl.CuocVCBo = oldVl.CuocVCBo;
                            vl.MaDinhMucBocXep = oldVl.MaDinhMucBocXep;
                            vl.PhamViBocXep = oldVl.PhamViBocXep;
                            vl.DmNCBocXep = oldVl.DmNCBocXep;
                            vl.DmMayBocXep = oldVl.DmMayBocXep;
                            vl.MaMayBocXep = oldVl.MaMayBocXep;
                            vl.MaDinhMucVCOTo = oldVl.MaDinhMucVCOTo;
                            vl.MaMayVCOTo = oldVl.MaMayVCOTo;
                            vl.MaDinhMucVCBo = oldVl.MaDinhMucVCBo;
                        }
                    }
                }

                // Load giá từ Bộ Đơn Giá vừa chọn (Ghi đè giá gốc từ thư viện chung)
                var boDonGiaRepo2 = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                var giaVl = boDonGiaRepo2.GetGiaVL(duToan.BoDonGiaId.Value).GroupBy(x => x.MaVL).ToDictionary(g => g.Key, g => g.First());
                var giaNc = boDonGiaRepo2.GetGiaNC(duToan.BoDonGiaId.Value).GroupBy(x => x.MaNC).ToDictionary(g => g.Key, g => g.First());
                var giaMay = boDonGiaRepo2.GetGiaMay(duToan.BoDonGiaId.Value).GroupBy(x => x.MaMay).ToDictionary(g => g.Key, g => g.First());

                foreach (var vl in duToan.BangTongHop.DanhSachVatLieu)
                {
                    if (giaVl.TryGetValue(vl.MaVatTu, out var g))
                    {
                        vl.GiaGoc = g.GiaGoc;
                        if (g.ChiPhiBocXep > 0 || g.CuocVCOTo > 0 || g.CuocVCBo > 0)
                        {
                            vl.ChiPhiBocXep = g.ChiPhiBocXep;
                            vl.CuocVCOTo = g.CuocVCOTo;
                            vl.CuocVCBo = g.CuocVCBo;
                        }
                        else if (vl.ChiPhiBocXep == 0 && vl.CuocVCOTo == 0 && vl.CuocVCBo == 0 && g.CuocVC > 0)
                        {
                            vl.CuocVanChuyen = g.CuocVC;
                        }
                    }
                }
                
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
                
                if (_tinhGiaForm != null && !_tinhGiaForm.IsDisposed)
                {
                    loading.Close();
                    loading.Dispose();
                    _tinhGiaForm.Activate();
                    return;
                }
                var form = new TinhGiaHienTruongForm(duToan, service, phanTichDonGiaService);
                _tinhGiaForm = form;
                form.FormClosed += (s, ev) => { _tinhGiaForm = null; };
                
                loading.Close();
                loading.Dispose();
                
                form.Show();
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
