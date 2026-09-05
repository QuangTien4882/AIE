using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.ExcelAddIn.Helpers;
using Newtonsoft.Json;

namespace AIE.ExcelAddIn.Forms
{
    public class ChonChiPhiThuVienForm : Form
    {
        private static List<ChiPhiKinhPhiItem> _danhSachThuVien = null;

        private DataGridView dgvThuVien;
        private TextBox txtTimKiem;
        private Label lblSoLuongChon;
        private Button btnThemChiPhiVaoThuVien;
        private Button btnXoaKhoiThuVien;
        private Button btnChonTatCa;
        private Button btnBoChonTatCa;
        private Button btnChon;
        private Button btnHuy;

        public List<ChiPhiKinhPhiItem> SelectedItems { get; private set; } = new List<ChiPhiKinhPhiItem>();

        public ChonChiPhiThuVienForm()
        {
            KhoiTaoThuVien();
            InitializeComponent();
            LoadData();

            this.Load += (s, e) =>
            {
                this.WindowState = FormWindowState.Maximized;
            };
            this.Shown += (s, e) =>
            {
                this.WindowState = FormWindowState.Maximized;
                txtTimKiem?.Focus();
            };
        }

        private static string GetStorageFilePath()
        {
            string appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AIE");
            if (!Directory.Exists(appData)) Directory.CreateDirectory(appData);
            return Path.Combine(appData, "thu_vien_chi_phi_custom.json");
        }

        private void KhoiTaoThuVien()
        {
            if (_danhSachThuVien != null) return;

            _danhSachThuVien = DinhMucTT38Engine.LayThuVienChiPhiChuan();

            // Nạp các khoản mục người dùng đã thêm trước đó
            try
            {
                string filePath = GetStorageFilePath();
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    var customItems = JsonConvert.DeserializeObject<List<ChiPhiKinhPhiItem>>(json);
                    if (customItems != null && customItems.Count > 0)
                    {
                        foreach (var ci in customItems)
                        {
                            if (!_danhSachThuVien.Any(x => x.MaChiPhi == ci.MaChiPhi))
                            {
                                ci.IsUserAdded = true;
                                _danhSachThuVien.Add(ci);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void LuuThuVienCustom()
        {
            try
            {
                string filePath = GetStorageFilePath();
                var customItems = _danhSachThuVien.Where(x => x.IsUserAdded).ToList();
                string json = JsonConvert.SerializeObject(customItems, Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
            catch { }
        }

        private void InitializeComponent()
        {
            this.Text = "Thư viện Khoản mục Chi phí chuẩn (Tư vấn đầu tư & Chi phí khác - Thông tư 38/2026/TT-BXD & BTC)";
            this.Size = new Size(1200, 750);
            this.MinimumSize = new Size(1000, 600);
            this.WindowState = FormWindowState.Maximized; // Yêu cầu 1: Mở mặc định 100% màn hình
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = UIHelper.GetFont(10.5f); // Yêu cầu 4: Font Be Vietnam Pro
            this.MinimizeBox = false;
            this.MaximizeBox = true;

            // =========================================================================
            // 1. TOP PANEL: THANH TÌM KIẾM & NÚT THAO TÁC THƯ VIỆN
            // =========================================================================
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 98,
                Padding = new Padding(12, 8, 12, 8),
                BackColor = Color.FromArgb(245, 248, 252)
            };

            var lblTieuDe = new Label
            {
                Text = "DANH MỤC THƯ VIỆN KHOẢN MỤC CHI PHÍ (TƯ VẤN ĐẦU TƯ & CHI PHÍ KHÁC THEO TT 38/2026 & BỘ TÀI CHÍNH)",
                Dock = DockStyle.Top,
                Height = 30,
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Container điều khiển tìm kiếm và nút chức năng
            var pnlToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoScroll = false,
                Padding = new Padding(0, 4, 0, 0)
            };

            var lblTim = new Label
            {
                Text = "🔍 Tìm kiếm:",
                AutoSize = true,
                Margin = new Padding(0, 8, 5, 0),
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102)
            };

            txtTimKiem = new TextBox
            {
                Width = 320,
                Height = 32,
                Font = UIHelper.GetFont(10.5f),
                Margin = new Padding(0, 4, 15, 0)
            };
            txtTimKiem.TextChanged += (s, e) => LocDanhSach();

            btnThemChiPhiVaoThuVien = new Button
            {
                Text = "➕ Thêm mục mới vào Thư viện",
                AutoSize = true,
                MinimumSize = new Size(240, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 2, 8, 0),
                Cursor = Cursors.Hand
            };
            btnThemChiPhiVaoThuVien.Click += BtnThemChiPhiVaoThuVien_Click;

            btnXoaKhoiThuVien = new Button
            {
                Text = "❌ Xóa mục khỏi Thư viện",
                AutoSize = true,
                MinimumSize = new Size(200, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                ForeColor = Color.DarkRed,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 2, 8, 0),
                Cursor = Cursors.Hand
            };
            btnXoaKhoiThuVien.Click += BtnXoaKhoiThuVien_Click;

            btnChonTatCa = new Button
            {
                Text = "☑ Chọn tất cả",
                AutoSize = true,
                MinimumSize = new Size(120, 36),
                Height = 36,
                Padding = new Padding(10, 0, 10, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 2, 6, 0),
                Cursor = Cursors.Hand
            };
            btnChonTatCa.Click += (s, e) => ChonTatCa(true);

            btnBoChonTatCa = new Button
            {
                Text = "☐ Bỏ chọn",
                AutoSize = true,
                MinimumSize = new Size(110, 36),
                Height = 36,
                Padding = new Padding(10, 0, 10, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 2, 0, 0),
                Cursor = Cursors.Hand
            };
            btnBoChonTatCa.Click += (s, e) => ChonTatCa(false);

            pnlToolbar.Controls.Add(lblTim);
            pnlToolbar.Controls.Add(txtTimKiem);
            pnlToolbar.Controls.Add(btnThemChiPhiVaoThuVien);
            pnlToolbar.Controls.Add(btnXoaKhoiThuVien);
            pnlToolbar.Controls.Add(btnChonTatCa);
            pnlToolbar.Controls.Add(btnBoChonTatCa);

            topPanel.Controls.Add(pnlToolbar);
            topPanel.Controls.Add(lblTieuDe);

            // =========================================================================
            // 2. BOTTOM PANEL: TỔNG KẾT & NÚT THỰC HIỆN
            // =========================================================================
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 65,
                Padding = new Padding(15, 12, 15, 12),
                BackColor = Color.FromArgb(245, 248, 252)
            };

            lblSoLuongChon = new Label
            {
                Text = "Đã chọn: 0 khoản mục",
                Dock = DockStyle.Left,
                Width = 350,
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var pnlBottomActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            btnChon = new Button
            {
                Text = "✔ Thêm các chi phí đã chọn vào Dự toán",
                AutoSize = true,
                MinimumSize = new Size(330, 40),
                Height = 40,
                Padding = new Padding(16, 0, 16, 0),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Margin = new Padding(0, 0, 12, 0),
                Cursor = Cursors.Hand
            };
            btnChon.Click += BtnChon_Click;

            btnHuy = new Button
            {
                Text = "Đóng",
                AutoSize = true,
                MinimumSize = new Size(110, 40),
                Height = 40,
                Padding = new Padding(16, 0, 16, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f),
                Margin = new Padding(0),
                Cursor = Cursors.Hand
            };
            btnHuy.Click += (s, e) => this.Close();

            pnlBottomActions.Controls.Add(btnChon);
            pnlBottomActions.Controls.Add(btnHuy);

            bottomPanel.Controls.Add(lblSoLuongChon);
            bottomPanel.Controls.Add(pnlBottomActions);

            // =========================================================================
            // 3. CENTER: BẢNG DỮ LIỆU THƯ VIỆN TỰ CO GIÃN 100%
            // =========================================================================
            dgvThuVien = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = UIHelper.GetFont(10.5f)
            };
            UIHelper.ApplyStyle(dgvThuVien);
            dgvThuVien.RowTemplate.Height = 36;
            dgvThuVien.CellValueChanged += (s, e) => CapNhatDemSoLuongChon();
            dgvThuVien.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvThuVien.IsCurrentCellDirty)
                    dgvThuVien.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            var colCheck = new DataGridViewCheckBoxColumn { Name = "colCheck", HeaderText = "Chọn", Width = 60, FillWeight = 5 };
            var colNhom = new DataGridViewTextBoxColumn { Name = "colNhom", HeaderText = "Nhóm chi phí", Width = 150, FillWeight = 14, ReadOnly = true };
            var colMa = new DataGridViewTextBoxColumn { Name = "colMa", HeaderText = "Mã hiệu", Width = 120, FillWeight = 12, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
            var colTen = new DataGridViewTextBoxColumn { Name = "colTen", HeaderText = "Tên khoản mục chi phí", Width = 400, FillWeight = 43, ReadOnly = true };
            var colCachTinh = new DataGridViewTextBoxColumn { Name = "colCachTinh", HeaderText = "Cách tính", Width = 120, FillWeight = 12, ReadOnly = true };
            var colTyLe = new DataGridViewTextBoxColumn { Name = "colTyLe", HeaderText = "Tỷ lệ %", Width = 80, FillWeight = 8, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
            var colVAT = new DataGridViewTextBoxColumn { Name = "colVAT", HeaderText = "VAT", Width = 75, FillWeight = 6, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };

            dgvThuVien.Columns.AddRange(colCheck, colNhom, colMa, colTen, colCachTinh, colTyLe, colVAT);

            this.Controls.Add(dgvThuVien);
            this.Controls.Add(bottomPanel);
            this.Controls.Add(topPanel);
        }

        private void LoadData(string tuKhoa = "")
        {
            dgvThuVien.Rows.Clear();
            var danhSach = _danhSachThuVien;

            if (!string.IsNullOrWhiteSpace(tuKhoa))
            {
                string tkLower = tuKhoa.Trim().ToLower();
                danhSach = _danhSachThuVien.Where(x =>
                    (x.TenChiPhi != null && x.TenChiPhi.ToLower().Contains(tkLower)) ||
                    (x.MaChiPhi != null && x.MaChiPhi.ToLower().Contains(tkLower)) ||
                    (x.GhiChuCachTinh != null && x.GhiChuCachTinh.ToLower().Contains(tkLower))
                ).ToList();
            }

            foreach (var item in danhSach)
            {
                string nhomStr = item.Nhom == NhomChiPhi.TuVanDauTuXD ? "IV. Tư vấn (G_TV)" : "V. Khác (G_K)";
                string cachTinhStr = item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc ? "Theo tỷ lệ %" : "Tự nhập tiền";
                string tyLeStr = item.TyLePhanTram > 0 ? UIHelper.FormatTyLe(item.TyLePhanTram) + "%" : "-";
                string vatStr = item.ThueSuatGTGT > 0 ? (item.ThueSuatGTGT * 100m).ToString("0.#", UIHelper.ViCulture) + "%" : "0% (Phí)";

                int rowIdx = dgvThuVien.Rows.Add(false, nhomStr, item.MaChiPhi, item.TenChiPhi, cachTinhStr, tyLeStr, vatStr);
                var row = dgvThuVien.Rows[rowIdx];
                row.Tag = item;

                if (item.IsUserAdded)
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 252, 235); // Tô màu nhẹ cho mục tự thêm
                }
            }

            CapNhatDemSoLuongChon();
        }

        private void LocDanhSach()
        {
            LoadData(txtTimKiem.Text);
        }

        private void ChonTatCa(bool chon)
        {
            for (int i = 0; i < dgvThuVien.Rows.Count; i++)
            {
                dgvThuVien.Rows[i].Cells["colCheck"].Value = chon;
            }
            CapNhatDemSoLuongChon();
        }

        private void CapNhatDemSoLuongChon()
        {
            int count = 0;
            for (int i = 0; i < dgvThuVien.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dgvThuVien.Rows[i].Cells["colCheck"].Value ?? false))
                    count++;
            }
            lblSoLuongChon.Text = $"Đã chọn: {count} / {dgvThuVien.Rows.Count} khoản mục chi phí";
        }

        private void BtnThemChiPhiVaoThuVien_Click(object sender, EventArgs e)
        {
            using var frm = new Form
            {
                Text = "Thêm khoản mục chi phí mới vào Thư viện",
                Size = new Size(620, 420),
                StartPosition = FormStartPosition.CenterParent,
                Font = UIHelper.GetFont(10.5f),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var pnlMain = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new Padding(20),
            };
            pnlMain.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            pnlMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // Tên khoản mục
            pnlMain.Controls.Add(new Label { Text = "Tên khoản mục chi phí:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Bold) }, 0, 0);
            var txtTen = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f) };
            pnlMain.Controls.Add(txtTen, 1, 0);

            // Nhóm chi phí
            pnlMain.Controls.Add(new Label { Text = "Nhóm chi phí:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 1);
            var cboNhom = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f) };
            cboNhom.Items.AddRange(new object[] { "Chi phí tư vấn đầu tư xây dựng (G_TV)", "Chi phí khác (G_K)" });
            cboNhom.SelectedIndex = 0;
            pnlMain.Controls.Add(cboNhom, 1, 1);

            // Cách tính
            pnlMain.Controls.Add(new Label { Text = "Cách tính:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 2);
            var cboCachTinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f) };
            cboCachTinh.Items.AddRange(new object[] { "Theo tỷ lệ %", "Tự nhập tiền trực tiếp" });
            cboCachTinh.SelectedIndex = 0;
            pnlMain.Controls.Add(cboCachTinh, 1, 2);

            // Tỷ lệ %
            pnlMain.Controls.Add(new Label { Text = "Tỷ lệ % (nếu tính %):", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 3);
            var txtTyLe = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), Text = "0,5" };
            pnlMain.Controls.Add(txtTyLe, 1, 3);

            // Thuế GTGT
            pnlMain.Controls.Add(new Label { Text = "Thuế suất GTGT:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 4);
            var cboVAT = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f) };
            cboVAT.Items.AddRange(new object[] { "10%", "8%", "5%", "0% (Phí / Lệ phí)" });
            cboVAT.SelectedIndex = 0;
            pnlMain.Controls.Add(cboVAT, 1, 4);

            // Buttons
            var pnlBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var btnSave = new Button
            {
                Text = "💾 Lưu vào Thư viện",
                Size = new Size(180, 38),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            var btnCancel = new Button
            {
                Text = "Hủy",
                Size = new Size(95, 38),
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };

            btnSave.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(txtTen.Text))
                {
                    MessageBox.Show("Vui lòng nhập tên khoản mục chi phí.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                decimal vatRate = 0.10m;
                if (cboVAT.SelectedIndex == 1) vatRate = 0.08m;
                else if (cboVAT.SelectedIndex == 2) vatRate = 0.05m;
                else if (cboVAT.SelectedIndex == 3) vatRate = 0m;

                var newItem = new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "CUSTOM_" + Guid.NewGuid().ToString("N").Substring(0, 6).ToUpper(),
                    TenChiPhi = txtTen.Text.Trim(),
                    Nhom = cboNhom.SelectedIndex == 0 ? NhomChiPhi.TuVanDauTuXD : NhomChiPhi.ChiPhiKhac,
                    CachTinh = cboCachTinh.SelectedIndex == 0 ? CachTinhChiPhi.TheoTyLeDinhMuc : CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    TyLePhanTram = UIHelper.ParseTyLe(txtTyLe.Text),
                    ThueSuatGTGT = vatRate,
                    KyHieu = cboNhom.SelectedIndex == 0 ? "Gtv_k" : "Gk",
                    IsActive = true,
                    IsUserAdded = true
                };

                _danhSachThuVien.Add(newItem);
                LuuThuVienCustom();
                LoadData(txtTimKiem.Text);
                frm.DialogResult = DialogResult.OK;
                frm.Close();
            };

            btnCancel.Click += (s, ev) => frm.Close();

            pnlBtns.Controls.Add(btnSave);
            pnlBtns.Controls.Add(btnCancel);
            pnlMain.Controls.Add(pnlBtns, 1, 5);

            frm.Controls.Add(pnlMain);
            frm.ShowDialog();
        }

        private void BtnXoaKhoiThuVien_Click(object sender, EventArgs e)
        {
            if (dgvThuVien.CurrentRow?.Tag is not ChiPhiKinhPhiItem item)
            {
                MessageBox.Show("Vui lòng chọn một khoản mục chi phí để xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"Bạn có chắc chắn muốn xóa khoản mục [{item.TenChiPhi}] khỏi Thư viện?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _danhSachThuVien.Remove(item);
                LuuThuVienCustom();
                LoadData(txtTimKiem.Text);
            }
        }

        private void BtnChon_Click(object sender, EventArgs e)
        {
            SelectedItems.Clear();
            for (int i = 0; i < dgvThuVien.Rows.Count; i++)
            {
                var row = dgvThuVien.Rows[i];
                bool isChecked = Convert.ToBoolean(row.Cells["colCheck"].Value ?? false);
                if (isChecked && row.Tag is ChiPhiKinhPhiItem item)
                {
                    var clone = item.Clone();
                    clone.IsActive = true;
                    clone.IsUserAdded = true;
                    SelectedItems.Add(clone);
                }
            }

            if (SelectedItems.Count == 0)
            {
                MessageBox.Show("Vui lòng tích chọn ít nhất một khoản mục chi phí.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
