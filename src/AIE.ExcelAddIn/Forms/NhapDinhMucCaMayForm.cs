using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Data.Repositories;

namespace AIE.ExcelAddIn.Forms;

public class NhapDinhMucCaMayForm : Form
{
    private readonly DinhMucCaMayRepository _repo;
    private readonly string _maMay;
    private readonly string _tenMay;
    private DinhMucCaMay_TT37 _dinhMuc;

    private TextBox txtNguyenGia;
    private TextBox txtKhauHao;
    private TextBox txtSuaChua;
    private TextBox txtChiPhiKhac;
    private TextBox txtXang;
    private TextBox txtDiezel;
    private TextBox txtDien;
    private TextBox txtNhanCong;
    private ComboBox cbNhomNhanCong;

    public NhapDinhMucCaMayForm(string maMay, string tenMay, DinhMucCaMayRepository repo)
    {
        _maMay = maMay;
        _tenMay = tenMay;
        _repo = repo;
        
        InitializeComponent();
        LoadData();
    }

    private void InitializeComponent()
    {
        this.Text = $"Định mức ca máy - {_maMay}";
        this.Size = new Size(580, 560);
        this.StartPosition = FormStartPosition.CenterParent;
        this.Font = new Font("Be Vietnam Pro", 9.5f);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };

        var lblTitle = new Label { Text = _tenMay, Font = new Font("Be Vietnam Pro", 11f, FontStyle.Bold), AutoSize = true, Location = new Point(20, 15), ForeColor = Color.FromArgb(0, 120, 215) };
        panel.Controls.Add(lblTitle);

        int y = 50;
        int lblWidth = 260; // Nới rộng để không bị che
        int txtWidth = 200;
        int gap = 38;

        txtNguyenGia = AddRow(panel, "Nguyên giá (đồng):", y, lblWidth, txtWidth); y += gap;
        txtKhauHao = AddRow(panel, "Tỷ lệ khấu hao (%/năm):", y, lblWidth, txtWidth); y += gap;
        txtSuaChua = AddRow(panel, "Tỷ lệ sửa chữa (%/năm):", y, lblWidth, txtWidth); y += gap;
        txtChiPhiKhac = AddRow(panel, "Chi phí khác (%/năm):", y, lblWidth, txtWidth); y += gap;
        
        var lblSep = new Label { Text = "--- Định mức tiêu hao ---", AutoSize = true, Location = new Point(20, y), ForeColor = Color.Gray };
        panel.Controls.Add(lblSep); y += 30;

        txtXang = AddRow(panel, "Xăng (lít/ca):", y, lblWidth, txtWidth); y += gap;
        txtDiezel = AddRow(panel, "Diezel (lít/ca):", y, lblWidth, txtWidth); y += gap;
        txtDien = AddRow(panel, "Điện (kWh/ca):", y, lblWidth, txtWidth); y += gap;
        txtNhanCong = AddRow(panel, "Nhân công điều khiển (người):", y, lblWidth, txtWidth); y += gap;

        var lblNhom = new Label { Text = "Nhóm thợ điều khiển:", AutoSize = true, Location = new Point(20, y + 4) };
        cbNhomNhanCong = new ComboBox { Location = new Point(lblWidth + 20, y), Width = txtWidth, DropDownStyle = ComboBoxStyle.DropDownList };
        cbNhomNhanCong.Items.AddRange(new object[] { "Nhóm 1", "Nhóm 2", "Nhóm 3", "Nhóm 4" });
        cbNhomNhanCong.SelectedIndex = 1; // Default nhóm 2
        panel.Controls.Add(lblNhom);
        panel.Controls.Add(cbNhomNhanCong);
        y += 50;

        var btnSave = new Button { Text = "💾 Lưu định mức", Location = new Point(200, y), Width = 150, Height = 40, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold) };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;
        panel.Controls.Add(btnSave);

        this.Controls.Add(panel);
    }

    private TextBox AddRow(Panel parent, string labelText, int y, int lblWidth, int txtWidth)
    {
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(20, y + 4) };
        var txt = new TextBox { Location = new Point(lblWidth + 20, y), Width = txtWidth, TextAlign = HorizontalAlignment.Right };
        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        return txt;
    }

    private void LoadData()
    {
        _dinhMuc = _repo.GetByMaMay(_maMay) ?? new DinhMucCaMay_TT37 { MaMay = _maMay };
        var viVn = new CultureInfo("vi-VN");
        txtNguyenGia.Text = _dinhMuc.NguyenGia == 0 ? "" : _dinhMuc.NguyenGia.ToString("N0", viVn);
        txtKhauHao.Text = _dinhMuc.KhauHao == 0 ? "" : _dinhMuc.KhauHao.ToString("N2", viVn);
        txtSuaChua.Text = _dinhMuc.SuaChua == 0 ? "" : _dinhMuc.SuaChua.ToString("N2", viVn);
        txtChiPhiKhac.Text = _dinhMuc.ChiPhiKhac == 0 ? "" : _dinhMuc.ChiPhiKhac.ToString("N2", viVn);
        
        txtXang.Text = _dinhMuc.DinhMucXang == 0 ? "" : _dinhMuc.DinhMucXang.ToString("N2", viVn);
        txtDiezel.Text = _dinhMuc.DinhMucDiezel == 0 ? "" : _dinhMuc.DinhMucDiezel.ToString("N2", viVn);
        txtDien.Text = _dinhMuc.DinhMucDien == 0 ? "" : _dinhMuc.DinhMucDien.ToString("N2", viVn);
        txtNhanCong.Text = _dinhMuc.SoLuongNhanCong == 0 ? "" : _dinhMuc.SoLuongNhanCong.ToString("N2", viVn);

        if (_dinhMuc.NhomNhanCong >= 1 && _dinhMuc.NhomNhanCong <= 4)
            cbNhomNhanCong.SelectedIndex = _dinhMuc.NhomNhanCong - 1;
            
        if (_dinhMuc.DinhMucXang > 0 || _dinhMuc.DinhMucDiezel > 0 || _dinhMuc.DinhMucDien > 0)
        {
            txtXang.Enabled = _dinhMuc.DinhMucXang > 0;
            txtDiezel.Enabled = _dinhMuc.DinhMucDiezel > 0;
            txtDien.Enabled = _dinhMuc.DinhMucDien > 0;
        }
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        try
        {
            var viVn = new CultureInfo("vi-VN");
            decimal ParseVn(string text)
            {
                if (string.IsNullOrWhiteSpace(text)) return 0;
                if (decimal.TryParse(text, NumberStyles.Any, viVn, out var val)) return val;
                return 0;
            }

            _dinhMuc.NguyenGia = ParseVn(txtNguyenGia.Text);
            _dinhMuc.KhauHao = ParseVn(txtKhauHao.Text);
            _dinhMuc.SuaChua = ParseVn(txtSuaChua.Text);
            _dinhMuc.ChiPhiKhac = ParseVn(txtChiPhiKhac.Text);
            _dinhMuc.DinhMucXang = ParseVn(txtXang.Text);
            _dinhMuc.DinhMucDiezel = ParseVn(txtDiezel.Text);
            _dinhMuc.DinhMucDien = ParseVn(txtDien.Text);
            _dinhMuc.SoLuongNhanCong = ParseVn(txtNhanCong.Text);
            _dinhMuc.NhomNhanCong = cbNhomNhanCong.SelectedIndex + 1;

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
