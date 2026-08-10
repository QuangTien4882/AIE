using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using AIE.Data;
using AIE.Data.ImportExport;

namespace AIE.ExcelAddIn.Forms
{
    public class ImportDatabaseForm : Form
    {
        private Label lblTitle;
        private Label lblStats;
        private Button btnImportDinhMuc;
        private Button btnImportDinhMucF1;
        private Button btnImportVatLieu;
        private Button btnImportNhanCong;
        private Button btnImportMay;
        private Button btnDong;
        private TextBox txtLog;

        private DatabaseManager _dbManager;

        public ImportDatabaseForm()
        {
            _dbManager = new DatabaseManager();
            InitializeComponents();
            RefreshStats();
        }

        private void InitializeComponents()
        {
            this.Text = "AIE - Nhập Database";
            this.Size = new Size(520, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Segoe UI", 9F);

            // Title
            lblTitle = new Label();
            lblTitle.Text = "NHẬP DỮ LIỆU VÀO DATABASE";
            lblTitle.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblTitle.Location = new Point(20, 15);
            lblTitle.AutoSize = true;
            this.Controls.Add(lblTitle);

            // Stats panel
            var pnlStats = new GroupBox();
            pnlStats.Text = "Thống kê Database hiện tại";
            pnlStats.Location = new Point(20, 50);
            pnlStats.Size = new Size(460, 70);
            this.Controls.Add(pnlStats);

            lblStats = new Label();
            lblStats.Location = new Point(15, 22);
            lblStats.Size = new Size(430, 40);
            lblStats.Text = "Đang tải...";
            pnlStats.Controls.Add(lblStats);

            // Buttons panel
            var pnlButtons = new GroupBox();
            pnlButtons.Text = "Chọn loại dữ liệu cần nhập";
            pnlButtons.Location = new Point(20, 130);
            pnlButtons.Size = new Size(460, 205);
            this.Controls.Add(pnlButtons);

            btnImportDinhMuc = CreateButton("1. Nhập Định Mức Công Tác (Từ file mẫu AIE)", 20, pnlButtons);
            btnImportDinhMuc.Click += BtnImportDinhMuc_Click;

            btnImportDinhMucF1 = CreateButton("1b. Nhập Định Mức Công Tác (Từ file F1 xuất ra)", 55, pnlButtons);
            btnImportDinhMucF1.Click += BtnImportDinhMucF1_Click;

            btnImportVatLieu = CreateButton("2. Nhập Giá Vật Liệu (Đà Nẵng)", 90, pnlButtons);
            btnImportVatLieu.Click += BtnImportVatLieu_Click;

            btnImportNhanCong = CreateButton("3. Nhập Giá Nhân Công (Đà Nẵng)", 125, pnlButtons);
            btnImportNhanCong.Click += BtnImportNhanCong_Click;

            btnImportMay = CreateButton("4. Nhập Giá Máy Thi Công (Đà Nẵng)", 160, pnlButtons);
            btnImportMay.Click += BtnImportMay_Click;

            // Log
            txtLog = new TextBox();
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.Location = new Point(20, 345);
            txtLog.Size = new Size(460, 85);
            txtLog.BackColor = Color.White;
            this.Controls.Add(txtLog);

            // Close button
            btnDong = new Button();
            btnDong.Text = "Đóng";
            btnDong.Size = new Size(100, 32);
            btnDong.Location = new Point(380, 440);
            btnDong.Click += (s, e) => this.Close();
            this.Controls.Add(btnDong);
        }

        private Button CreateButton(string text, int top, Control parent)
        {
            var btn = new Button();
            btn.Text = text;
            btn.Size = new Size(420, 30);
            btn.Location = new Point(20, top);
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 215);
            btn.Cursor = Cursors.Hand;
            parent.Controls.Add(btn);
            return btn;
        }

        private void RefreshStats()
        {
            try
            {
                var stats = _dbManager.GetStats();
                lblStats.Text = stats.TomTat;
            }
            catch (Exception ex)
            {
                lblStats.Text = "Lỗi khi đọc DB: " + ex.Message;
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }

        private string[] ChooseExcelFiles(string title)
        {
            var ofd = new OpenFileDialog();
            ofd.Title = title;
            ofd.Filter = "Excel Files|*.xlsx;*.xls";
            ofd.FilterIndex = 1;
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
                return ofd.FileNames;
            return null;
        }

        private void BtnImportDinhMuc_Click(object sender, EventArgs e)
        {
            var filePaths = ChooseExcelFiles("Chọn file 1_DinhMucCongTac.xlsx");
            if (filePaths == null || filePaths.Length == 0) return;

            foreach (var filePath in filePaths)
            {
                try
                {
                    Log($"Đang import Định Mức Công Tác từ {Path.GetFileName(filePath)}...");
                    var result = _dbManager.ImportDinhMucCongTac(filePath);
                    Log("Hoàn thành! " + result.TomTat);
                    foreach (var err in result.DanhSachLoi)
                        Log("  ⚠ " + err);
                }
                catch (Exception ex)
                {
                    Log("LỖI: " + ex.Message);
                    MessageBox.Show($"Lỗi import file {Path.GetFileName(filePath)}: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            RefreshStats();
        }

        private void BtnImportDinhMucF1_Click(object sender, EventArgs e)
        {
            var filePaths = ChooseExcelFiles("Chọn file định mức xuất từ phần mềm F1");
            if (filePaths == null || filePaths.Length == 0) return;

            foreach (var filePath in filePaths)
            {
                try
                {
                    Log($"Đang phân tích file F1: {Path.GetFileName(filePath)}...");
                    var result = _dbManager.ImportF1DinhMucCongTac(filePath);
                    Log("Hoàn thành! " + result.TomTat);
                }
                catch (Exception ex)
                {
                    Log("LỖI: " + ex.Message);
                    MessageBox.Show($"Lỗi import F1 file {Path.GetFileName(filePath)}: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            RefreshStats();
        }

        private void BtnImportVatLieu_Click(object sender, EventArgs e)
        {
            var filePaths = ChooseExcelFiles("Chọn file 2_GiaVatLieu.xlsx");
            if (filePaths == null || filePaths.Length == 0) return;

            foreach (var filePath in filePaths)
            {
                try
                {
                    Log($"Đang import Giá Vật Liệu từ {Path.GetFileName(filePath)}...");
                    var result = _dbManager.ImportGiaVatLieu(filePath);
                    Log("Hoàn thành! " + result.TomTat);
                    foreach (var err in result.DanhSachLoi)
                        Log("  ⚠ " + err);
                }
                catch (Exception ex)
                {
                    Log("LỖI: " + ex.Message);
                    MessageBox.Show($"Lỗi import file {Path.GetFileName(filePath)}: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            RefreshStats();
        }

        private void BtnImportNhanCong_Click(object sender, EventArgs e)
        {
            var filePaths = ChooseExcelFiles("Chọn file 3_GiaNhanCong.xlsx");
            if (filePaths == null || filePaths.Length == 0) return;

            foreach (var filePath in filePaths)
            {
                try
                {
                    Log($"Đang import Giá Nhân Công từ {Path.GetFileName(filePath)}...");
                    var result = _dbManager.ImportGiaNhanCong(filePath);
                    Log("Hoàn thành! " + result.TomTat);
                    foreach (var err in result.DanhSachLoi)
                        Log("  ⚠ " + err);
                }
                catch (Exception ex)
                {
                    Log("LỖI: " + ex.Message);
                    MessageBox.Show($"Lỗi import file {Path.GetFileName(filePath)}: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            RefreshStats();
        }

        private void BtnImportMay_Click(object sender, EventArgs e)
        {
            var filePaths = ChooseExcelFiles("Chọn file 4_GiaMayThiCong.xlsx");
            if (filePaths == null || filePaths.Length == 0) return;

            foreach (var filePath in filePaths)
            {
                try
                {
                    Log($"Đang import Giá Máy Thi Công từ {Path.GetFileName(filePath)}...");
                    var result = _dbManager.ImportGiaMayThiCong(filePath);
                    Log("Hoàn thành! " + result.TomTat);
                    foreach (var err in result.DanhSachLoi)
                        Log("  ⚠ " + err);
                }
                catch (Exception ex)
                {
                    Log("LỖI: " + ex.Message);
                    MessageBox.Show($"Lỗi import file {Path.GetFileName(filePath)}: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            RefreshStats();
        }
    }
}
