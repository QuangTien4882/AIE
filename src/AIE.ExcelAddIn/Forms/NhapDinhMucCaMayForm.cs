using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms;

public class NhapDinhMucCaMayForm : Form
{
    private readonly DinhMucCaMayRepository _repo;
    private readonly string _maMay;
    private readonly string _tenMay;
    private DinhMucCaMay_TT37 _dinhMuc;

    private TextBox txtNguyenGia;
    private TextBox txtSoCaNam;
    private TextBox txtKhauHao;
    private TextBox txtSuaChua;
    private TextBox txtChiPhiKhac;
    
    private TextBox txtXang;
    private TextBox txtDiezel;
    private TextBox txtDien;

    private DataGridView dgvNhanCong;
    private BindingList<WorkerRow> _workerList;

    /// <summary>
    /// Danh sách chức danh nhân công vận hành, điều khiển máy theo bảng công bố.
    /// Key = mã chức danh, Value = (Tên hiển thị, Nhóm nhân công vận hành II-x)
    /// </summary>
    public static readonly Dictionary<string, (string TenHienThi, int NhomVanHanh)> DanhSachChucDanh = new()
    {
        // Nhóm II-1: Nhân công vận hành máy, điều khiển máy
        { "VH",   ("Nhân công vận hành máy", 1) },
        // Nhóm II-2: Lái xe
        { "LX",   ("Lái xe", 2) },
        // Nhóm II-3: Thủy thủ, thợ máy, thợ điện
        { "TT",   ("Thủy thủ", 3) },
        { "TM",   ("Thợ máy", 3) },
        { "TDIEN",("Thợ điện", 3) },
        // Nhóm II-4: Máy trưởng, điện trưởng, thuyền trưởng, thuyền phó...
        { "MTR",  ("Máy trưởng", 4) },
        { "MAY2", ("Máy II", 4) },
        { "DTRU", ("Điện trưởng", 4) },
        { "TTRU", ("Thuyền trưởng", 4) },
        { "TPHO", ("Thuyền phó", 4) },
        { "KTV1", ("Kỹ thuật viên cuốc I", 4) },
        { "KTV2", ("Kỹ thuật viên cuốc II", 4) },
        { "MAY1", ("Máy I", 4) },
    };

    public class WorkerRow
    {
        public string ChucDanh { get; set; } = "LX";
        public decimal SoLuong { get; set; }
    }

    public NhapDinhMucCaMayForm(string maMay, string tenMay, DinhMucCaMayRepository repo)
    {
        _maMay = maMay;
        _tenMay = tenMay;
        _repo = repo;
        
        InitializeComponent();
        AIE.ExcelAddIn.Helpers.FormStateHelper.Attach(this);
        LoadData();
        this.Shown += (s, e) => dgvNhanCong?.AutoFit();
    }

    private void InitializeComponent()
    {
        this.Text = $"Định mức ca máy - {_maMay}";
        this.StartPosition = FormStartPosition.CenterParent;
        this.Font = new Font("Be Vietnam Pro", 9.5f);
        this.FormBorderStyle = FormBorderStyle.Sizable;
        
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size(750, Math.Min(800, (int)(workingArea.Height * 0.9)));
        this.MinimumSize = new Size(720, 550);
        this.MaximizeBox = true;
        this.MinimizeBox = true;

        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20), AutoScroll = true };
        this.Controls.Add(panel);

        var lblTitle = new Label { Text = _tenMay, Font = new Font("Be Vietnam Pro", 11f, FontStyle.Bold), AutoSize = true, Location = new Point(20, 15), ForeColor = Color.FromArgb(0, 120, 215) };
        panel.Controls.Add(lblTitle);

        int y = 50;
        int lblWidth = 240; 
        int txtWidth = 200;
        int gap = 35;

        // Nhóm 1: Thông tin chung
        var gbChung = new GroupBox { Text = "Thông tin chung", Location = new Point(20, y), Width = 690, Height = 210, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        int gY = 25;
        txtNguyenGia = AddRow(gbChung, "Nguyên giá (đồng):", gY, lblWidth, txtWidth); gY += gap;
        txtSoCaNam = AddRow(gbChung, "Số ca/năm:", gY, lblWidth, txtWidth); gY += gap;
        txtKhauHao = AddRow(gbChung, "Tỷ lệ khấu hao (%/năm):", gY, lblWidth, txtWidth); gY += gap;
        txtSuaChua = AddRow(gbChung, "Tỷ lệ sửa chữa (%/năm):", gY, lblWidth, txtWidth); gY += gap;
        txtChiPhiKhac = AddRow(gbChung, "Chi phí khác (%/năm):", gY, lblWidth, txtWidth); gY += gap;
        panel.Controls.Add(gbChung);
        y += gbChung.Height + 15;

        // Nhóm 2: Nhiên liệu
        var gbNhienLieu = new GroupBox { Text = "Định mức tiêu hao nhiên liệu", Location = new Point(20, y), Width = 690, Height = 140, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
        gY = 25;
        txtXang = AddRow(gbNhienLieu, "Xăng (lít/ca):", gY, lblWidth, txtWidth); gY += gap;
        txtDiezel = AddRow(gbNhienLieu, "Diezel (lít/ca):", gY, lblWidth, txtWidth); gY += gap;
        txtDien = AddRow(gbNhienLieu, "Điện (kWh/ca):", gY, lblWidth, txtWidth); gY += gap;
        panel.Controls.Add(gbNhienLieu);
        y += gbNhienLieu.Height + 15;

        // Nhóm 3: Nhân công vận hành
        var gbNhanCong = new GroupBox { Text = "Nhân công vận hành", Location = new Point(20, y), Width = 690, Height = 170, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
        var btnAddTho = new Button { Text = "➕ Thêm thợ", Location = new Point(560, 20), Width = 110, Height = 30, BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        btnAddTho.FlatAppearance.BorderSize = 0;
        btnAddTho.Click += BtnAddTho_Click;
        gbNhanCong.Controls.Add(btnAddTho);

        dgvNhanCong = new DataGridView
        {
            Location = new Point(20, 60),
            Width = 650,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            BackgroundColor = Color.White,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        
        var colChucDanh = new DataGridViewComboBoxColumn
        {
            Name = "ChucDanh",
            HeaderText = "Nhóm thợ (Vận hành)",
            DataPropertyName = "ChucDanh",
            DataSource = DanhSachChucDanh.Select(x => new { Key = x.Key, Display = x.Value.TenHienThi }).ToList(),
            DisplayMember = "Display",
            ValueMember = "Key",
            ValueType = typeof(string),
            FlatStyle = FlatStyle.Flat
        };
        dgvNhanCong.Columns.Add(colChucDanh);

        dgvNhanCong.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoLuong", HeaderText = "Số lượng", DataPropertyName = "SoLuong", DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##", Alignment = DataGridViewContentAlignment.MiddleRight } });
        
        var colDelete = new DataGridViewButtonColumn { Name = "Delete", HeaderText = "Xóa", Text = "X", UseColumnTextForButtonValue = true, Width = 50 };
        dgvNhanCong.Columns.Add(colDelete);
        dgvNhanCong.CellContentClick += DgvNhanCong_CellContentClick;
        AIE.ExcelAddIn.Helpers.UIHelper.ApplyStyle(dgvNhanCong);
        dgvNhanCong.RowHeadersVisible = false; // Override back to false
        dgvNhanCong.AllowUserToResizeRows = false;
        gbNhanCong.Controls.Add(dgvNhanCong);
        
        gbNhanCong.Height = 170; // temporary height
        panel.Controls.Add(gbNhanCong);

        var btnSave = new Button { Text = "💾 Lưu định mức", Width = 150, Height = 40, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;
        panel.Controls.Add(btnSave);
        
        // Handle layout when form is shown/resized
        this.Load += (s, e) => {
            gbNhanCong.Height = this.ClientSize.Height - gbNhanCong.Top - 70;
            dgvNhanCong.Height = gbNhanCong.Height - 60 - 20;
            btnSave.Location = new Point((this.ClientSize.Width - btnSave.Width) / 2, this.ClientSize.Height - 60);
            
            // Now apply anchors so they resize automatically
            gbNhanCong.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnSave.Anchor = AnchorStyles.Bottom;
        };
    }

    private TextBox AddRow(Control parent, string labelText, int y, int lblWidth, int txtWidth)
    {
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(20, y + 4) };
        var txt = new TextBox { Location = new Point(lblWidth + 20, y), Width = txtWidth, TextAlign = HorizontalAlignment.Right };
        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        return txt;
    }

    private void BtnAddTho_Click(object? sender, EventArgs e)
    {
        _workerList.Add(new WorkerRow { ChucDanh = "LX", SoLuong = 1 });
    }

    private void DgvNhanCong_CellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0 && e.ColumnIndex == dgvNhanCong.Columns["Delete"].Index)
        {
            _workerList.RemoveAt(e.RowIndex);
        }
    }

    private void LoadData()
    {
        _dinhMuc = _repo.GetByMaMay(_maMay) ?? new DinhMucCaMay_TT37 { MaMay = _maMay };
        var viVn = new CultureInfo("vi-VN");
        txtNguyenGia.Text = _dinhMuc.NguyenGia == 0 ? "" : _dinhMuc.NguyenGia.ToString("N0", viVn);
        txtSoCaNam.Text = _dinhMuc.SoCaNam == 0 ? "250" : _dinhMuc.SoCaNam.ToString();
        txtKhauHao.Text = _dinhMuc.KhauHao == 0 ? "" : _dinhMuc.KhauHao.ToString("N2", viVn);
        txtSuaChua.Text = _dinhMuc.SuaChua == 0 ? "" : _dinhMuc.SuaChua.ToString("N2", viVn);
        txtChiPhiKhac.Text = _dinhMuc.ChiPhiKhac == 0 ? "" : _dinhMuc.ChiPhiKhac.ToString("N2", viVn);
        
        txtXang.Text = _dinhMuc.DinhMucXang == 0 ? "" : _dinhMuc.DinhMucXang.ToString("N2", viVn);
        txtDiezel.Text = _dinhMuc.DinhMucDiezel == 0 ? "" : _dinhMuc.DinhMucDiezel.ToString("N2", viVn);
        txtDien.Text = _dinhMuc.DinhMucDien == 0 ? "" : _dinhMuc.DinhMucDien.ToString("N2", viVn);
        
        _workerList = new BindingList<WorkerRow>();
        
        // Parse ThanhPhanNhanCong trực tiếp để hỗ trợ cả format cũ (nhom số) và mới (mã chức danh)
        if (!string.IsNullOrWhiteSpace(_dinhMuc.ThanhPhanNhanCong))
        {
            foreach (var part in _dinhMuc.ThanhPhanNhanCong.Split(';'))
            {
                var kv = part.Trim().Split(':');
                if (kv.Length == 2 && decimal.TryParse(kv[1], NumberStyles.Any, CultureInfo.InvariantCulture, out decimal sl))
                {
                    string key = kv[0].Trim();
                    // Nếu key là số thuần (format cũ), chuyển sang mã chức danh tương ứng
                    if (int.TryParse(key, out int nhomCu))
                    {
                        key = nhomCu switch
                        {
                            1 => "VH",
                            2 => "LX",
                            3 => "TT",
                            4 or 5 or 6 => "MTR",
                            _ => "VH"
                        };
                    }
                    // Đảm bảo key hợp lệ
                    if (!DanhSachChucDanh.ContainsKey(key)) key = "VH";
                    _workerList.Add(new WorkerRow { ChucDanh = key, SoLuong = sl });
                }
            }
        }
        else if (_dinhMuc.NhomNhanCong > 0 && _dinhMuc.SoLuongNhanCong > 0)
        {
            // Legacy fallback
            string key = _dinhMuc.NhomNhanCong switch
            {
                1 => "VH",
                2 => "LX",
                3 => "TT",
                _ => "MTR"
            };
            _workerList.Add(new WorkerRow { ChucDanh = key, SoLuong = _dinhMuc.SoLuongNhanCong });
        }
        
        dgvNhanCong.DataSource = _workerList;
        dgvNhanCong?.AutoFit();
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        try
        {
            dgvNhanCong.EndEdit();
            
            var viVn = new CultureInfo("vi-VN");
            decimal ParseVn(string text)
            {
                if (string.IsNullOrWhiteSpace(text)) return 0;
                if (decimal.TryParse(text, NumberStyles.Any, viVn, out var val)) return val;
                return 0;
            }

            _dinhMuc.NguyenGia = ParseVn(txtNguyenGia.Text);
            _dinhMuc.SoCaNam = int.TryParse(txtSoCaNam.Text, out var sc) ? sc : 250;
            _dinhMuc.KhauHao = ParseVn(txtKhauHao.Text);
            _dinhMuc.SuaChua = ParseVn(txtSuaChua.Text);
            _dinhMuc.ChiPhiKhac = ParseVn(txtChiPhiKhac.Text);
            _dinhMuc.DinhMucXang = ParseVn(txtXang.Text);
            _dinhMuc.DinhMucDiezel = ParseVn(txtDiezel.Text);
            _dinhMuc.DinhMucDien = ParseVn(txtDien.Text);
            
            // Build ThanhPhanNhanCong string
            var parts = new List<string>();
            var descParts = new List<string>();
            foreach (var w in _workerList)
            {
                if (w.SoLuong > 0)
                {
                    int nhomSo = DanhSachChucDanh.ContainsKey(w.ChucDanh) ? DanhSachChucDanh[w.ChucDanh].NhomVanHanh : 1;
                    string tenCD = DanhSachChucDanh.ContainsKey(w.ChucDanh) ? DanhSachChucDanh[w.ChucDanh].TenHienThi : w.ChucDanh;
                    // Rút gọn tên chức danh cho chuỗi hiển thị
                    string tenNgan = tenCD.Split(',')[0].Trim().ToLower();
                    
                    parts.Add($"{w.ChucDanh}:{w.SoLuong.ToString("0.##", CultureInfo.InvariantCulture)}");
                    descParts.Add($"{w.SoLuong.ToString("0.##", viVn)} {tenNgan}");
                }
            }
            
            _dinhMuc.ThanhPhanNhanCong = string.Join(";", parts);
            
            _dinhMuc.NhanCongString = string.Join(" + ", descParts);
            
            // Xóa dữ liệu cũ
            _dinhMuc.SoLuongNhanCong = 0;
            _dinhMuc.NhomNhanCong = 0;

            _repo.Upsert(_dinhMuc);
            MessageBox.Show("Lưu định mức thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
