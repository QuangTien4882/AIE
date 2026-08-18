using AIE.Core.Models;
using AIE.ExcelAddIn.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

public class TinhGiaHienTruongForm : Form
{
    private readonly DuToan _duToan;
    private readonly BangTongHopVatTu _bangTongHop;
    private readonly PhanTichVatTuService _service;
    private readonly PhanTichDonGiaService _donGiaService;
    private static readonly CultureInfo ViVn = new CultureInfo("vi-VN");
    private static readonly string SettingsDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
    private static readonly string SettingsFile = Path.Combine(SettingsDir, "column_widths.cfg");
    
    private TabControl tabControl;
    private DataGridView dgvVL, dgvNC, dgvMay;
    
    private TextBox txtGiaXang, txtGiaDiezel, txtGiaDien;
    private bool _suppressRecalc = false;

    public TinhGiaHienTruongForm(DuToan duToan, PhanTichVatTuService service, PhanTichDonGiaService donGiaService)
    {
        _duToan = duToan;
        _bangTongHop = duToan.BangTongHop;
        _service = service;
        _donGiaService = donGiaService;
        
        InitializeComponent();
        LoadData();
        RestoreColumnWidths();
        
        this.FormClosing += (s, e) => SaveColumnWidths();
    }

    private void InitializeComponent()
    {
        this.Text = "Tính Giá vật tư hiện trường";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        var topPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(12) };
        var lblTitle = new Label { Text = "BẢNG TỔNG HỢP VÀ TÍNH GIÁ VẬT TƯ", Font = new Font("Be Vietnam Pro", 14f, FontStyle.Bold), AutoSize = true, Location = new Point(12, 16), ForeColor = Color.FromArgb(0, 120, 215) };
        
        var rightPanel = new FlowLayoutPanel 
        { 
            Dock = DockStyle.Right, 
            FlowDirection = FlowDirection.RightToLeft, 
            Width = 450,
            WrapContents = false,
            Padding = new Padding(0, 0, 0, 0)
        };

        var btnThoat = new Button 
        { 
            Text = "Đóng", 
            Size = new Size(100, 36), 
            BackColor = Color.White, 
            ForeColor = Color.Black, 
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Margin = new Padding(5, 0, 0, 0)
        };
        btnThoat.FlatAppearance.BorderColor = Color.LightGray;
        btnThoat.Click += (s, e) => this.Close();

        var btnApGia = new Button 
        { 
            Text = "Áp giá vào Dự toán", 
            Size = new Size(180, 36), 
            BackColor = Color.FromArgb(40, 167, 69), 
            ForeColor = Color.White, 
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Margin = new Padding(5, 0, 0, 0)
        };
        btnApGia.FlatAppearance.BorderSize = 0;
        btnApGia.Click += BtnApGia_Click;

        var btnLuu = new Button 
        { 
            Text = "Lưu đơn giá", 
            Size = new Size(140, 36), 
            BackColor = Color.FromArgb(0, 120, 215), 
            ForeColor = Color.White, 
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Margin = new Padding(5, 0, 0, 0)
        };
        btnLuu.FlatAppearance.BorderSize = 0;
        btnLuu.Click += BtnLuu_Click;

        rightPanel.Controls.Add(btnThoat);
        rightPanel.Controls.Add(btnApGia);
        rightPanel.Controls.Add(btnLuu);

        topPanel.Controls.Add(lblTitle);
        topPanel.Controls.Add(rightPanel);

        tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold), DrawMode = TabDrawMode.OwnerDrawFixed };
        tabControl.DrawItem += TabControl_DrawItem;

        // Tab Vật liệu
        var tabVL = new TabPage("Vật liệu");
        dgvVL = CreateGrid("dgvVL");
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã VL", DataPropertyName = "MaVatTu", ReadOnly = true, Width = 100 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên vật liệu", DataPropertyName = "TenVatTu", ReadOnly = true, Width = 300 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TongKhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "TongKhoiLuong", ReadOnly = true, Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaGoc", HeaderText = "Giá gốc", DataPropertyName = "GiaGoc", ReadOnly = false, Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "CuocVanChuyen", HeaderText = "Cước VC", DataPropertyName = "CuocVanChuyen", Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá hiện trường", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvVL.Font, FontStyle.Bold) } });
        dgvVL.CellValueChanged += DgvVL_CellValueChanged;
        dgvVL.CellParsing += DgvVL_CellParsing;
        tabVL.Controls.Add(dgvVL);

        // Tab Nhân công
        var tabNC = new TabPage("Nhân công");
        dgvNC = CreateGrid("dgvNC");
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã NC", DataPropertyName = "MaVatTu", ReadOnly = true, Width = 100 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên nhân công", DataPropertyName = "TenVatTu", ReadOnly = true, Width = 300 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TongKhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "TongKhoiLuong", ReadOnly = true, Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá nhân công", DataPropertyName = "GiaHienTruong", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200), ForeColor = Color.Red, Font = new Font(dgvNC.Font, FontStyle.Bold) } });
        tabNC.Controls.Add(dgvNC);

        // Tab Máy thi công
        var tabMay = new TabPage("Máy thi công");
        
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(10), BackColor = Color.FromArgb(240, 248, 255) };
        int px = 10;
        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px);
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px);
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px);

        dgvMay = CreateGrid("dgvMay");
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã Máy", DataPropertyName = "MaVatTu", ReadOnly = true, Width = 100 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên máy", DataPropertyName = "TenVatTu", ReadOnly = true, Width = 250 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "TongKhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "TongKhoiLuong", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N2", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhauHao", HeaderText = "CP Khấu hao", DataPropertyName = "ChiPhiKhauHao", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiSuaChua", HeaderText = "CP Sửa chữa", DataPropertyName = "ChiPhiSuaChua", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhac", HeaderText = "CP Khác", DataPropertyName = "ChiPhiKhac", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhiemLieu", HeaderText = "CP Nhiên liệu", DataPropertyName = "ChiPhiNhiemLieu", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhanCong", HeaderText = "CP Thợ lái", DataPropertyName = "ChiPhiNhanCong", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 140, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });
        
        tabMay.Controls.Add(dgvMay);
        tabMay.Controls.Add(fuelPanel);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);

        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
    }

    private void BtnLuu_Click(object sender, EventArgs e)
    {
        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            _service.SaveGia(_bangTongHop);
            MessageBox.Show("Đã lưu đơn giá hiện tại vào cơ sở dữ liệu!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi lưu đơn giá: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnApGia_Click(object sender, EventArgs e)
    {
        try
        {
            // 1. Force kết thúc edit ở các grid để lấy giá trị mới nhất
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            // 2. Tính đơn giá chi tiết
            _donGiaService.TinhDonGiaChiTiet(_duToan);

            // 3. Lưu giá trị hiện tại vào Bộ đơn giá (nếu có)
            if (_duToan.BoDonGiaId.HasValue)
            {
                var db = new AIE.Data.DatabaseManager();
                var boDonGiaRepo = new AIE.Data.Repositories.BoDonGiaRepository(db.Context);
                int bId = _duToan.BoDonGiaId.Value;

                foreach (var vl in _duToan.BangTongHop.DanhSachVatLieu)
                    boDonGiaRepo.SaveGiaVL(bId, vl.MaVatTu, vl.GiaGoc, vl.CuocVanChuyen, vl.GiaHienTruong);

                foreach (var nc in _duToan.BangTongHop.DanhSachNhanCong)
                    boDonGiaRepo.SaveGiaNC(bId, nc.MaVatTu, nc.GiaGoc); // Giá gốc của nhân công là đơn giá

                foreach (var may in _duToan.BangTongHop.DanhSachMay)
                    boDonGiaRepo.SaveGiaMay(bId, may.MaVatTu, may.GiaGoc); // Giá gốc của máy được tính và cập nhật

                // Cập nhật giá nhiên liệu vào BoDonGia
                if (decimal.TryParse(txtGiaXang.Text, out decimal xang) &&
                    decimal.TryParse(txtGiaDiezel.Text, out decimal diezel) &&
                    decimal.TryParse(txtGiaDien.Text, out decimal dien))
                {
                    boDonGiaRepo.UpdateFuelPrices(bId, xang, diezel, dien);
                    
                    // Update BangTongHop
                    _duToan.BangTongHop.GiaXang = xang;
                    _duToan.BangTongHop.GiaDiezel = diezel;
                    _duToan.BangTongHop.GiaDien = dien;
                }
            }
            else
            {
                // Fallback: save to JSON if no BoDonGia (for old projects without a set)
                try
                {
                    var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                    if (!System.IO.Directory.Exists(settingsDir)) System.IO.Directory.CreateDirectory(settingsDir);
                    var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                    
                    decimal.TryParse(txtGiaXang.Text, out decimal xang);
                    decimal.TryParse(txtGiaDiezel.Text, out decimal diezel);
                    decimal.TryParse(txtGiaDien.Text, out decimal dien);

                    var obj = new { GiaXang = xang, GiaDiezel = diezel, GiaDien = dien };
                    System.IO.File.WriteAllText(fuelFile, Newtonsoft.Json.JsonConvert.SerializeObject(obj));
                    
                    _duToan.BangTongHop.GiaXang = xang;
                    _duToan.BangTongHop.GiaDiezel = diezel;
                    _duToan.BangTongHop.GiaDien = dien;
                }
                catch { }
            }

            // 4. Ghi vào Excel
            var excelService = new LapDuToanExcelService();
            excelService.WriteDonGiaToExcel(_duToan);

            MessageBox.Show("Đã áp giá thành công vào file Excel!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi áp giá: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        var g = e.Graphics;
        var page = tabControl.TabPages[e.Index];
        var tabBounds = tabControl.GetTabRect(e.Index);
        
        if (e.State == DrawItemState.Selected)
        {
            g.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), tabBounds);
            TextRenderer.DrawText(g, page.Text, page.Font, tabBounds, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        else
        {
            g.FillRectangle(new SolidBrush(Color.FromArgb(240, 240, 240)), tabBounds);
            TextRenderer.DrawText(g, page.Text, page.Font, tabBounds, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    private void DgvVL_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
    {
        if (e.ColumnIndex == dgvVL.Columns["CuocVanChuyen"].Index || e.ColumnIndex == dgvVL.Columns["GiaGoc"].Index)
        {
            if (e.Value != null)
            {
                var text = e.Value.ToString().Replace(".", "").Replace(",", ".");
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsedValue))
                {
                    e.Value = parsedValue;
                    e.ParsingApplied = true;
                }
            }
        }
    }

    private void DgvVL_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex >= 0 && (e.ColumnIndex == dgvVL.Columns["CuocVanChuyen"].Index || e.ColumnIndex == dgvVL.Columns["GiaGoc"].Index))
        {
            dgvVL.InvalidateRow(e.RowIndex);
        }
    }

    private TextBox AddFuelInput(Panel parent, string labelText, ref int x)
    {
        int labelWidth = 130;
        var lbl = new Label { Text = labelText, AutoSize = false, Size = new Size(labelWidth, 25), Location = new Point(x, 24), Font = new Font("Be Vietnam Pro", 9.5f) };
        var txt = new TextBox { Location = new Point(x + labelWidth, 20), Width = 100, Font = new Font("Be Vietnam Pro", 10f), TextAlign = HorizontalAlignment.Right };
        
        // Khi click vào ô: xóa dấu chấm để sẵn sàng nhập số mới
        txt.Enter += (s, e) => 
        { 
            _suppressRecalc = true;
            if (txt.Text == "0") 
                txt.Clear(); 
            else 
                txt.Text = txt.Text.Replace(".", ""); 
            txt.SelectAll();
            _suppressRecalc = false;
        };
        
        // Khi rời ô: format lại chuẩn Việt Nam và tính toán lại
        txt.Leave += (s, e) => 
        { 
            _suppressRecalc = true;
            if (decimal.TryParse(txt.Text, out var val)) 
                txt.Text = val.ToString("N0", ViVn); 
            _suppressRecalc = false;
            RecalculateMachinePrices();
        };

        // Chỉ cho phép nhập số
        txt.KeyPress += (s, e) =>
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
                e.Handled = true;
        };

        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        x += labelWidth + 100 + 30; // space for next
        return txt;
    }

    private void LoadData()
    {
        _suppressRecalc = true;
        txtGiaXang.Text = _bangTongHop.GiaXang.ToString("N0", ViVn);
        txtGiaDiezel.Text = _bangTongHop.GiaDiezel.ToString("N0", ViVn);
        txtGiaDien.Text = _bangTongHop.GiaDien.ToString("N0", ViVn);
        _suppressRecalc = false;

        dgvVL.DataSource = new BindingSource { DataSource = _bangTongHop.DanhSachVatLieu };
        dgvNC.DataSource = new BindingSource { DataSource = _bangTongHop.DanhSachNhanCong };
        dgvMay.DataSource = new BindingSource { DataSource = _bangTongHop.DanhSachMay };
    }

    private void RecalculateMachinePrices()
    {
        if (_suppressRecalc) return;
        
        decimal ParseVn(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            // Xóa dấu chấm phân cách hàng ngàn, giữ lại số
            var clean = text.Replace(".", "");
            if (decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) return val;
            return 0;
        }

        _bangTongHop.GiaXang = ParseVn(txtGiaXang.Text);
        _bangTongHop.GiaDiezel = ParseVn(txtGiaDiezel.Text);
        _bangTongHop.GiaDien = ParseVn(txtGiaDien.Text);

        _service.TinhGiaMayThiCong(_bangTongHop);
        dgvMay.Refresh();
    }

    private DataGridView CreateGrid(string gridName)
    {
        return new DataGridView
        {
            Name = gridName,
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
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
            RowTemplate = { Height = 35 },
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, // Không tự co giãn
            ScrollBars = ScrollBars.Both,
            EditMode = DataGridViewEditMode.EditOnEnter
        };
    }

    // === Lưu / Khôi phục độ rộng cột ===
    
    private void SaveColumnWidths()
    {
        try
        {
            if (!Directory.Exists(SettingsDir)) Directory.CreateDirectory(SettingsDir);
            
            var lines = new List<string>();
            SaveGridWidths(lines, dgvVL);
            SaveGridWidths(lines, dgvNC);
            SaveGridWidths(lines, dgvMay);
            File.WriteAllLines(SettingsFile, lines);
        }
        catch { /* Ignore save errors */ }
    }

    private void SaveGridWidths(List<string> lines, DataGridView dgv)
    {
        foreach (DataGridViewColumn col in dgv.Columns)
        {
            lines.Add($"{dgv.Name}|{col.Name}|{col.Width}");
        }
    }

    private void RestoreColumnWidths()
    {
        try
        {
            if (!File.Exists(SettingsFile)) return;
            
            var lines = File.ReadAllLines(SettingsFile);
            var grids = new Dictionary<string, DataGridView>
            {
                { dgvVL.Name, dgvVL },
                { dgvNC.Name, dgvNC },
                { dgvMay.Name, dgvMay }
            };
            
            foreach (var line in lines)
            {
                var parts = line.Split('|');
                if (parts.Length != 3) continue;
                
                if (grids.TryGetValue(parts[0], out var dgv))
                {
                    if (dgv.Columns.Contains(parts[1]) && int.TryParse(parts[2], out var width))
                    {
                        dgv.Columns[parts[1]].Width = width;
                    }
                }
            }
        }
        catch { /* Ignore restore errors */ }
    }
}
