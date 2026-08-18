using AIE.Core.Models;
using AIE.Data;
using AIE.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

public class QuanLyDonGiaForm : Form
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

    private TabControl tabControl;
    private TextBox txtSearch;
    private Label lblCount;

    // Data
    private VatLieuRepository _vlRepo;
    private NhanCongRepository _ncRepo;
    private MayThiCongRepository _mayRepo;
    private DinhMucCaMayRepository _mayDmRepo;

    private List<VatLieu> _allVL;
    private List<NhanCong> _allNC;
    private List<MayThiCong> _allMay;

    private DataGridView dgvVL, dgvNC, dgvMay;

    public QuanLyDonGiaForm(VatLieuRepository vlRepo, NhanCongRepository ncRepo, MayThiCongRepository mayRepo, DinhMucCaMayRepository mayDmRepo)
    {
        _vlRepo = vlRepo;
        _ncRepo = ncRepo;
        _mayRepo = mayRepo;
        _mayDmRepo = mayDmRepo;

        _allVL = _vlRepo.GetAll().ToList();
        _allNC = _ncRepo.GetAll().ToList();
        _allMay = _mayRepo.GetAll().ToList();

        InitializeComponent();
        LoadData();
    }

    private void InitializeComponent()
    {
        this.Text = "Quản Lý Đơn Giá";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MinimumSize = new Size(800, 500);
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        // ===== Top Panel: Search =====
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(12, 10, 12, 5) };

        var lblSearch = new Label { Text = "🔍 Tìm kiếm:", AutoSize = true, Location = new Point(12, 16) };
        txtSearch = new TextBox { Location = new Point(120, 12), Width = 350, Font = new Font("Be Vietnam Pro", 11f) };
        // Set placeholder via Win32 cue banner (PlaceholderText not available in .NET 4.8)
        SendMessage(txtSearch.Handle, 0x1501, 1, "Gõ tên hoặc mã để lọc...");
        txtSearch.TextChanged += TxtSearch_TextChanged;

        var lblHint = new Label { Text = "💡 Click đúp vào ô Đơn giá để chỉnh sửa", AutoSize = true, Location = new Point(490, 16), ForeColor = Color.FromArgb(0, 120, 215), Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Italic) };

        lblCount = new Label { AutoSize = true, ForeColor = Color.Gray, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        lblCount.Location = new Point(850, 16);

        topPanel.Controls.Add(lblSearch);
        topPanel.Controls.Add(txtSearch);
        topPanel.Controls.Add(lblHint);
        topPanel.Controls.Add(lblCount);

        topPanel.Resize += (s, e) => {
            lblCount.Location = new Point(topPanel.Width - lblCount.Width - 20, 16);
        };

        // ===== Tab Control =====
        tabControl = new TabControl { Dock = DockStyle.Fill };
        tabControl.Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold);

        var tabVL = new TabPage("Vật liệu");
        var tabNC = new TabPage("Nhân công");
        var tabMay = new TabPage("Máy thi công");

        dgvVL = CreateGrid();
        dgvNC = CreateGrid();
        dgvMay = CreateGrid();

        tabVL.Controls.Add(dgvVL);
        tabNC.Controls.Add(dgvNC);
        tabMay.Controls.Add(dgvMay);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);
        tabControl.SelectedIndexChanged += (s, e) => ApplyFilter();

        // ===== Bottom Panel: Buttons =====
        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(12, 8, 12, 8) };

        var btnSave = new Button
        {
            Text = "💾  Lưu thay đổi",
            Width = 160, Height = 38,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            Anchor = AnchorStyles.Right
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;
        btnSave.Location = new Point(bottomPanel.Width - 185, 8);

        var btnClose = new Button
        {
            Text = "Đóng",
            Width = 100, Height = 38,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10f),
            Anchor = AnchorStyles.Right
        };
        btnClose.Click += (s, e) => this.Close();
        btnClose.Location = new Point(bottomPanel.Width - 300, 8);

        bottomPanel.Controls.Add(btnSave);
        bottomPanel.Controls.Add(btnClose);

        // ===== Layout =====
        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
        this.Controls.Add(bottomPanel);

        // Fix button anchoring after layout
        this.Load += (s, e) =>
        {
            btnSave.Location = new Point(bottomPanel.Width - 185, 8);
            btnClose.Location = new Point(bottomPanel.Width - 300, 8);
        };
        this.Resize += (s, e) =>
        {
            btnSave.Location = new Point(bottomPanel.Width - 185, 8);
            btnClose.Location = new Point(bottomPanel.Width - 300, 8);
        };
    }

    private DataGridView CreateGrid()
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            RowHeadersVisible = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Regular),
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(245, 248, 252) },
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(0, 5, 0, 5)
            },
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeight = 40,
            RowTemplate = { Height = 32 },
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            EditMode = DataGridViewEditMode.EditOnEnter
        };

        dgv.CellFormatting += Dgv_CellFormatting;
        return dgv;
    }

    private void LoadData()
    {
        // === VL ===
        dgvVL.Columns.Clear();
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVL", HeaderText = "Mã VL", DataPropertyName = "MaVL", ReadOnly = true, FillWeight = 15 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVL", HeaderText = "Tên vật liệu", DataPropertyName = "TenVL", ReadOnly = true, FillWeight = 40 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 10 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGia", HeaderText = "Đơn giá", DataPropertyName = "DonGia", FillWeight = 18, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", DataPropertyName = "GhiChu", FillWeight = 17 });

        // === NC ===
        dgvNC.Columns.Clear();
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaNC", HeaderText = "Mã NC", DataPropertyName = "MaNC", ReadOnly = true, FillWeight = 15 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenNC", HeaderText = "Tên nhân công", DataPropertyName = "TenNC", ReadOnly = true, FillWeight = 40 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 10 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGia", HeaderText = "Đơn giá", DataPropertyName = "DonGia", FillWeight = 18, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", DataPropertyName = "GhiChu", FillWeight = 17 });

        // === May ===
        dgvMay.Columns.Clear();
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaMay", HeaderText = "Mã máy", DataPropertyName = "MaMay", ReadOnly = true, FillWeight = 15 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenMay", HeaderText = "Tên máy thi công", DataPropertyName = "TenMay", ReadOnly = true, FillWeight = 40 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 10 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGia", HeaderText = "Đơn giá", DataPropertyName = "DonGia", FillWeight = 18, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", DataPropertyName = "GhiChu", FillWeight = 17 });

        var menuMay = new ContextMenuStrip();
        var itemNhap = new ToolStripMenuItem("📝 Nhập Định Mức Ca Máy (TT37/2026)");
        itemNhap.Click += (s, e) => {
            if (dgvMay.SelectedRows.Count > 0)
            {
                var row = dgvMay.SelectedRows[0];
                var maMay = row.Cells["MaMay"].Value?.ToString();
                var tenMay = row.Cells["TenMay"].Value?.ToString();
                if (maMay != null && tenMay != null)
                {
                    var oldDm = _mayDmRepo.GetByMaMay(maMay);
                    using var form = new NhapDinhMucCaMayForm(maMay, tenMay, _mayDmRepo);
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        var newDm = _mayDmRepo.GetByMaMay(maMay);
                        if (oldDm != null && newDm != null)
                        {
                            var oldCP = oldDm.NguyenGia * (oldDm.KhauHao + oldDm.SuaChua + oldDm.ChiPhiKhac) / 100m / 250m;
                            var newCP = newDm.NguyenGia * (newDm.KhauHao + newDm.SuaChua + newDm.ChiPhiKhac) / 100m / 250m;
                            
                            // Delta for NhanCong
                            var oldNC = _ncRepo.GetAll().FirstOrDefault(x => x.Nhom == oldDm.NhomNhanCong)?.DonGia ?? 0;
                            var newNC = _ncRepo.GetAll().FirstOrDefault(x => x.Nhom == newDm.NhomNhanCong)?.DonGia ?? 0;
                            var diffNC = (newDm.SoLuongNhanCong * newNC) - (oldDm.SoLuongNhanCong * oldNC);

                            var diffTotal = (newCP - oldCP) + diffNC;

                            var mayItem = _allMay.FirstOrDefault(m => m.MaMay == maMay);
                            if (mayItem != null)
                            {
                                mayItem.DonGia += diffTotal;
                                _mayRepo.Upsert(mayItem);
                                dgvMay.Refresh();
                            }
                        }
                    }
                }
            }
        };
        menuMay.Items.Add(itemNhap);
        dgvMay.ContextMenuStrip = menuMay;

        dgvMay.CellMouseDown += (s, e) => {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgvMay.ClearSelection();
                dgvMay.Rows[e.RowIndex].Selected = true;
                dgvMay.CurrentCell = dgvMay.Rows[e.RowIndex].Cells[e.ColumnIndex];
            }
        };

        ApplyFilter();
    }

    private void TxtSearch_TextChanged(object sender, EventArgs e)
    {
        ApplyFilter();
    }

    private void ApplyFilter()
    {
        string keyword = txtSearch.Text.Trim().ToLower();

        var filteredVL = string.IsNullOrEmpty(keyword)
            ? _allVL
            : _allVL.Where(x => MatchAllKeywords(keyword, x.MaVL, x.TenVL)).ToList();
        dgvVL.DataSource = new BindingSource { DataSource = filteredVL };

        var filteredNC = string.IsNullOrEmpty(keyword)
            ? _allNC
            : _allNC.Where(x => MatchAllKeywords(keyword, x.MaNC, x.TenNC)).ToList();
        dgvNC.DataSource = new BindingSource { DataSource = filteredNC };

        var filteredMay = string.IsNullOrEmpty(keyword)
            ? _allMay
            : _allMay.Where(x => MatchAllKeywords(keyword, x.MaMay, x.TenMay)).ToList();
        dgvMay.DataSource = new BindingSource { DataSource = filteredMay };

        switch (tabControl.SelectedIndex)
        {
            case 0: lblCount.Text = $"Hiển thị {filteredVL.Count} / {_allVL.Count} vật liệu"; break;
            case 1: lblCount.Text = $"Hiển thị {filteredNC.Count} / {_allNC.Count} nhân công"; break;
            case 2: lblCount.Text = $"Hiển thị {filteredMay.Count} / {_allMay.Count} máy thi công"; break;
        }
    }

    /// <summary>
    /// Tìm kiếm thông minh: tách từ khoá thành các từ rời, yêu cầu TẤT CẢ các từ đều
    /// xuất hiện trong chuỗi ghép, không phân biệt thứ tự.
    /// </summary>
    private static bool MatchAllKeywords(string keyword, params string[] fields)
    {
        var combined = string.Join(" ", fields.Where(f => f != null)).ToLower();
        var words = keyword.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return words.All(w => combined.Contains(w));
    }

    private void BtnSave_Click(object sender, EventArgs e)
    {
        try
        {
            int countUpdated = 0;

            foreach (var vl in _allVL)
            {
                _vlRepo.Upsert(vl);
                countUpdated++;
            }
            foreach (var nc in _allNC)
            {
                _ncRepo.Upsert(nc);
                countUpdated++;
            }
            foreach (var may in _allMay)
            {
                _mayRepo.Upsert(may);
                countUpdated++;
            }

            MessageBox.Show($"Đã lưu thành công {countUpdated} bản ghi!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
    {
        var dgv = (DataGridView)sender;
        if (dgv.Columns[e.ColumnIndex].Name == "DonGia" && e.Value != null)
        {
            if (decimal.TryParse(e.Value.ToString(), out decimal val) && val == 0)
            {
                e.CellStyle.ForeColor = Color.LightGray;
                e.CellStyle.Font = new Font(dgv.Font, FontStyle.Italic);
            }
        }
    }
}
