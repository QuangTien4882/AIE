using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AIE.Data;
using AIE.Data.ImportExport;
using AIE.ExcelAddIn.Forms;
using System.IO;
using System;

namespace AIE.ExcelAddIn.Ribbon
{
    [ComVisible(true)]
    public class AieRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabAIE' label='AIE Dự Toán'>
                    
                    <group id='groupThietLap' label='Thiết lập dữ liệu'>
                      <button id='btnImport' showLabel='false' screentip='Nhập Database' size='large' imageMso='DatabaseInsert' onAction='OnImportClicked' />
                      <button id='btnTraCuu' showLabel='false' screentip='Tra Cứu Định Mức' size='large' imageMso='Search' onAction='OnTraCuuClicked' />
                      <button id='btnTaoTemplate' showLabel='false' screentip='Tạo File Mẫu' size='large' imageMso='FileSaveAs' onAction='OnTaoTemplateClicked' />
                      <button id='btnDeleteDb' showLabel='false' screentip='Xóa toàn bộ Database' size='large' imageMso='RecordsDeleteRecord' onAction='OnDeleteDbClicked' />
                    </group>

                    <group id='groupThamDinh' label='Thẩm định dự toán'>
                      <button id='btnKiemTra' showLabel='false' screentip='Kiểm Tra' size='large' imageMso='ReviewAcceptChange' onAction='OnKiemTraClicked' />
                      <button id='btnBaoCaoTD' showLabel='false' screentip='Xuất Báo Cáo' size='large' imageMso='ExportExcel' onAction='OnBaoCaoTDClicked' />
                    </group>

                    <group id='groupLapDuToan' label='Lập dự toán'>
                      <button id='btnGoiDonGia' showLabel='false' screentip='Gọi Đơn Giá' size='large' imageMso='DollarSign' onAction='OnGoiDonGiaClicked' />
                      <button id='btnTinhTongHop' showLabel='false' screentip='Tính Tổng Hợp' size='large' imageMso='CalculateNow' onAction='OnTinhTongHopClicked' />
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

        public void OnTraCuuClicked(IRibbonControl control)
        {
            try
            {
                var form = new TraCuuDinhMucForm();
                form.ShowDialog();
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
                MessageBox.Show("Lỗi khi tạo file mẫu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        MessageBox.Show("Không tìm thấy dữ liệu công tác nào. Vui lòng kiểm tra lại cấu hình cột.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 2. Thẩm định
                    var db = new DatabaseManager(); // Đảm bảo đã init db
                    var repo = new AIE.Data.Repositories.CongTacRepository(db.Context);
                    var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
                    var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                    var mayRepo = new AIE.Data.Repositories.MayThiCongRepository(db.Context);
                    var engine = new AIE.ExcelAddIn.Services.ThamDinhEngine(repo, vlRepo, ncRepo, mayRepo);
                    
                    var ketQua = engine.KiemTra(danhSachCongTac);

                    // 3. Xuất kết quả
                    var writer = new AIE.ExcelAddIn.Services.ThamDinhExcelWriter();
                    writer.ExportResult(config, ketQua);
                    
                    int soLoi = ketQua.Sum(x => x.DanhSachSaiLech.Count);
                    MessageBox.Show($"Thẩm định hoàn tất!\nĐã kiểm tra: {ketQua.Count} công tác.\nPhát hiện: {soLoi} sai lệch.\nKết quả đã được xuất ra sheet 'KQ_ThamDinh'.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra thẩm định: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnBaoCaoTDClicked(IRibbonControl control)
        {
            MessageBox.Show("Chức năng Xuất báo cáo thẩm định đang được phát triển.", "AIE Dự Toán");
        }

        public void OnGoiDonGiaClicked(IRibbonControl control)
        {
            MessageBox.Show("Chức năng Gọi đơn giá chi tiết đang được phát triển.", "AIE Dự Toán");
        }

        public void OnTinhTongHopClicked(IRibbonControl control)
        {
            MessageBox.Show("Chức năng Tính tổng hợp dự toán đang được phát triển.", "AIE Dự Toán");
        }
    }
}
