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
    }

    private void InitializeComponent()
    {
        this.Text = "Tính Giá vật tư hiện trường";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MinimumSize = new Size(800, 500);
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        // ===== Top Panel: Title & Tools =====
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(12, 10, 12, 5) };
        var lblTitle = new Label { Text = "BẢNG TỔNG HỢP VÀ TÍNH GIÁ VẬT TƯ", Font = new Font("Be Vietnam Pro", 14f, FontStyle.Bold), AutoSize = true, Location = new Point(12, 12), ForeColor = Color.FromArgb(0, 120, 215) };
        topPanel.Controls.Add(lblTitle);

        var btnAutoFit = new Button
        {
            Text = "↕ Tự động giãn cột/dòng",
            AutoSize = true,
            Location = new Point(topPanel.Width - 190, 10),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(240, 240, 240),
            Font = new Font("Be Vietnam Pro", 9f)
        };
        btnAutoFit.FlatAppearance.BorderColor = Color.LightGray;
        btnAutoFit.Click += (s, e) =>
        {
            DataGridView activeGrid = null;
            if (tabControl.SelectedTab == tabControl.TabPages[0]) activeGrid = dgvVL;
            else if (tabControl.SelectedTab == tabControl.TabPages[1]) activeGrid = dgvNC;
            else if (tabControl.SelectedTab == tabControl.TabPages[2]) activeGrid = dgvMay;

            if (activeGrid != null)
            {
                activeGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                activeGrid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            }
        };
        topPanel.Controls.Add(btnAutoFit);

        // ===== Bottom Panel: Buttons =====
        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 55, Padding = new Padding(12, 8, 12, 8) };

        var btnTable = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 1
        };
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
        btnTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));

        var btnThoat = new Button { Text = "Đóng", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnThoat.FlatAppearance.BorderColor = Color.LightGray;
        btnThoat.Click += (s, e) => this.Close();

        var btnApGia = new Button { Text = "📋  Áp giá vào Dự toán", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnApGia.FlatAppearance.BorderSize = 0;
        btnApGia.Click += BtnApGia_Click;

        var btnCapNhatGoc = new Button { Text = "🔄  Cập nhật vào đơn giá gốc", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.FromArgb(255, 152, 0), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnCapNhatGoc.FlatAppearance.BorderSize = 0;
        btnCapNhatGoc.Click += BtnCapNhatGoc_Click;

        var btnLuu = new Button { Text = "💾  Lưu đơn giá", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 5, 0), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnLuu.FlatAppearance.BorderSize = 0;
        btnLuu.Click += BtnLuu_Click;

        btnTable.Controls.Add(btnThoat, 0, 0);
        btnTable.Controls.Add(btnApGia, 1, 0);
        btnTable.Controls.Add(btnCapNhatGoc, 2, 0);
        btnTable.Controls.Add(btnLuu, 3, 0);

        bottomPanel.Controls.Add(btnTable);

        // ===== Tab Control =====
        tabControl = new TabControl { Dock = DockStyle.Fill };
        tabControl.Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold);
        AIE.ExcelAddIn.Helpers.UIHelper.ApplyStyle(tabControl);

        var tabVL = new TabPage("Vật liệu");
        var tabNC = new TabPage("Nhân công");
        var tabMay = new TabPage("Máy thi công");

        // === Tab Vật liệu ===
        dgvVL = CreateGrid("dgvVL");
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã VL", DataPropertyName = "MaVatTu", ReadOnly = true, FillWeight = 15 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên vật liệu", DataPropertyName = "TenVatTu", ReadOnly = true, FillWeight = 40 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 8, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TongKhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "TongKhoiLuong", ReadOnly = true, FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0.00", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaGoc", HeaderText = "Giá gốc", DataPropertyName = "GiaGoc", ReadOnly = false, FillWeight = 14, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "CuocVanChuyen", HeaderText = "Cước VC", DataPropertyName = "CuocVanChuyen", FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá hiện trường", DataPropertyName = "GiaHienTruong", ReadOnly = true, FillWeight = 16, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red } });
        dgvVL.CellValueChanged += DgvVL_CellValueChanged;
        dgvVL.CellParsing += DgvVL_CellParsing;
        tabVL.Controls.Add(dgvVL);

        // === Tab Nhân công ===
        dgvNC = CreateGrid("dgvNC");
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã NC", DataPropertyName = "MaVatTu", ReadOnly = true, FillWeight = 15 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên nhân công", DataPropertyName = "TenVatTu", ReadOnly = true, FillWeight = 40 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 10, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TongKhoiLuong", HeaderText = "Khối lượng", DataPropertyName = "TongKhoiLuong", ReadOnly = true, FillWeight = 15, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0.00", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá nhân công", DataPropertyName = "GiaHienTruong", ReadOnly = false, FillWeight = 20, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        tabNC.Controls.Add(dgvNC);

        // === Tab Máy thi công (cột giống QuanLyDonGiaForm) ===
        var mayContainer = new Panel { Dock = DockStyle.Fill };
        
        // --- Fuel Panel: Vùng (readonly) + Giá nhiên liệu ---
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8, 6, 8, 6) };
        fuelPanel.BackColor = Color.FromArgb(240, 248, 255);

        int px = 10;

        // Vùng áp dụng (readonly label, not editable)
        var lblVungLabel = new Label { Text = "Vùng áp dụng:", AutoSize = true, Location = new Point(px, 10), Font = new Font("Be Vietnam Pro", 9f) };
        fuelPanel.Controls.Add(lblVungLabel);
        px += lblVungLabel.PreferredWidth + 4;

        string vungText = _duToan.VungApDung switch
        {
            AIE.Core.Enums.Vung.VungII => "Vùng II",
            AIE.Core.Enums.Vung.VungIII => "Vùng III",
            AIE.Core.Enums.Vung.VungIV => "Vùng IV",
            AIE.Core.Enums.Vung.CuLaoCham => "Cù Lao Chàm",
            _ => "Vùng II"
        };
        var lblVungValue = new Label
        {
            Text = vungText,
            AutoSize = true,
            Location = new Point(px, 10),
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215)
        };
        fuelPanel.Controls.Add(lblVungValue);
        px += lblVungValue.PreferredWidth + 30;

        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px);
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px);
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px);

        dgvMay = CreateGrid("dgvMay");
        dgvMay.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        // Col 0
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVatTu", HeaderText = "Mã hiệu", DataPropertyName = "MaVatTu", ReadOnly = true, Width = 90, Frozen = true });
        // Col 1
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVatTu", HeaderText = "Tên máy và thiết bị", DataPropertyName = "TenVatTu", ReadOnly = true, Width = 200, Frozen = true, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
        // Col 2
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoCaNam", HeaderText = "Số ca/năm", DataPropertyName = "SoCaNam", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        // Col 3 (Định mức start)
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DM_KhauHao", HeaderText = "Khấu hao", DataPropertyName = "DmKhauHao", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        // Col 4
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DM_SuaChua", HeaderText = "Sửa chữa", DataPropertyName = "DmSuaChua", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        // Col 5 (Định mức end)
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DM_Khac", HeaderText = "Khác", DataPropertyName = "DmKhac", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        // Col 6
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucNhienLieuDisplay", HeaderText = "Định mức tiêu hao NL", DataPropertyName = "DinhMucNhienLieuDisplay", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, WrapMode = DataGridViewTriState.True } });
        // Col 7
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "HeSoNhienLieuPhu", HeaderText = "Hệ số NL phụ", DataPropertyName = "HeSoNhienLieuPhu", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##", Alignment = DataGridViewContentAlignment.MiddleCenter } });
        // Col 8
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhanCongVanHanhDisplay", HeaderText = "Nhân công vận hành", DataPropertyName = "NhanCongVanHanhDisplay", ReadOnly = true, Width = 200, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True } });
        // Col 9
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 10 (Chi phí start)
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhauHao", HeaderText = "Khấu hao", DataPropertyName = "ChiPhiKhauHao", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 11
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiSuaChua", HeaderText = "Sửa chữa", DataPropertyName = "ChiPhiSuaChua", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 12 (Chi phí end)
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhac", HeaderText = "Khác", DataPropertyName = "ChiPhiKhac", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 13
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhiemLieu", HeaderText = "CP Nhiên liệu", DataPropertyName = "ChiPhiNhiemLieu", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 14
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhanCong", HeaderText = "Lương thợ", DataPropertyName = "ChiPhiNhanCong", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        // Col 15
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 130, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });

        dgvMay.CellPainting += DgvMay_CellPainting;
        dgvMay.CellDoubleClick += DgvMay_CellDoubleClick;

        mayContainer.Controls.Add(dgvMay);
        mayContainer.Controls.Add(fuelPanel);
        tabMay.Controls.Add(mayContainer);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);

        // ===== Layout =====
        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
        this.Controls.Add(bottomPanel);
    }

    private void BtnCapNhatGoc_Click(object sender, EventArgs e)
    {
        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            var db = new AIE.Data.DatabaseManager();
            var vlRepo = new AIE.Data.Repositories.VatLieuRepository(db.Context);
            var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

            // Cập nhật giá vật liệu
            foreach (var vl in _bangTongHop.DanhSachVatLieu)
            {
                var vlMaster = vlRepo.GetByMa(vl.MaVatTu);
                if (vlMaster != null)
                {
                    vlMaster.DonGia = vl.GiaGoc;
                    vlMaster.CuocVanChuyen = vl.CuocVanChuyen;
                    vlRepo.Upsert(vlMaster);
                }
            }

            // Cập nhật giá nhân công
            foreach (var nc in _bangTongHop.DanhSachNhanCong)
            {
                var ncMaster = ncRepo.GetByMa(nc.MaVatTu);
                if (ncMaster != null)
                {
                    ncMaster.SetDonGia(_duToan.VungApDung, nc.GiaHienTruong);
                    ncRepo.Upsert(ncMaster);
                }
            }

            MessageBox.Show("Đã cập nhật thành công vào bộ đơn giá gốc!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi cập nhật: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnLuu_Click(object sender, EventArgs e)
    {
        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            _service.UpdateMasterDatabase(_bangTongHop);
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
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(x, 10), Font = new Font("Be Vietnam Pro", 9f) };
        parent.Controls.Add(lbl);
        x += lbl.PreferredWidth + 4;
        var txt = new TextBox { Location = new Point(x, 7), Width = 90, Font = new Font("Be Vietnam Pro", 9.5f), TextAlign = HorizontalAlignment.Right };
        
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

        parent.Controls.Add(txt);
        x += 100; // space for next
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

    private void DgvMay_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        // Định mức: cols 3, 4, 5  |  Chi phí: cols 10, 11, 12
        Helpers.GridHelper.PaintMergedHeader(sender, e, dgvMay, 3, 5, "Định mức", 10, 12, "Chi phí");
    }

    private void DgvMay_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var mayHT = _bangTongHop.DanhSachMay[e.RowIndex];
        if (mayHT?.DinhMuc != null)
        {
            var dbManager = new AIE.Data.DatabaseManager();
            var mayDmRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(dbManager.Context);
            using var frm = new NhapDinhMucCaMayForm(mayHT.MaVatTu, mayHT.TenVatTu ?? "", mayDmRepo);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                // Reload the DinhMuc from DB after editing
                var updatedDm = mayDmRepo.GetByMaMay(mayHT.MaVatTu);
                if (updatedDm != null)
                {
                    mayHT.DinhMuc = updatedDm;
                }
                RecalculateMachinePrices();
            }
        }
    }

    private DataGridView CreateGrid(string gridName)
    {
        var dgv = new DataGridView
        {
            Name = gridName,
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            RowHeadersVisible = true,
            RowHeadersWidth = 25,
            AllowUserToResizeRows = true,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Regular),
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
            EditMode = DataGridViewEditMode.EditOnEnter
        };

        AIE.ExcelAddIn.Helpers.UIHelper.ApplyStyle(dgv);
        
        return dgv;
    }
}
