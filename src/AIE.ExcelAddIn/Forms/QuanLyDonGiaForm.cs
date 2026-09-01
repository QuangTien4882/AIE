using AIE.Core.Models;
using AIE.Data;
using AIE.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using AIE.ExcelAddIn.Helpers;

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
    private List<DonGiaMayViewModel> _mayViewModels;
    private AIE.Core.Enums.Vung _vungApDung = AIE.Core.Enums.Vung.VungII;

    private DataGridView dgvVL, dgvNC, dgvMay;
    
    // Fuel price inputs for Máy thi công
    private TextBox txtGiaXang, txtGiaDiezel, txtGiaDien;
    private bool _suppressRecalc = false;

    private static readonly System.Globalization.CultureInfo ViVn = new System.Globalization.CultureInfo("vi-VN");

    public QuanLyDonGiaForm(VatLieuRepository vlRepo, NhanCongRepository ncRepo, MayThiCongRepository mayRepo, DinhMucCaMayRepository mayDmRepo)
    {
        _vlRepo = vlRepo;
        _ncRepo = ncRepo;
        _mayRepo = mayRepo;
        _mayDmRepo = mayDmRepo;

        _allVL = _vlRepo.GetAll().ToList();
        
        var dbManager = new AIE.Data.DatabaseManager();
        dbManager.SeedNhanCong();
        
        _allNC = _ncRepo.GetAll().ToList();
        _allMay = _mayRepo.GetAll().ToList();
        
        BuildMayViewModels();

        InitializeComponent();
        LoadData();
    }

    private void BuildMayViewModels()
    {
        _mayViewModels = new List<DonGiaMayViewModel>();
        foreach (var may in _allMay)
        {
            var dm = _mayDmRepo.GetByMaMay(may.MaMay);
            var vm = new DonGiaMayViewModel
            {
                MaHieu = may.MaMay,
                Ten = may.TenMay,
                DonVi = may.DonVi,
                MayThiCong = may
            };

            if (dm != null)
            {
                vm.NguyenGia = dm.NguyenGia;
                vm.SoCaNam = dm.SoCaNam > 0 ? dm.SoCaNam : 250;
                vm.TyLeKhauHao = dm.KhauHao;
                vm.TyLeSuaChua = dm.SuaChua;
                vm.TyLeKhac = dm.ChiPhiKhac;
                vm.DinhMucXang = dm.DinhMucXang;
                vm.DinhMucDiezel = dm.DinhMucDiezel;
                vm.DinhMucDien = dm.DinhMucDien;
                vm.SoLuongNhanCong = dm.SoLuongNhanCong;
                vm.NhomNhanCong = dm.NhomNhanCong;

                // Calculate cost components
                vm.ChiPhiKhauHao = dm.NguyenGia * dm.KhauHao / 100m / vm.SoCaNam;
                vm.ChiPhiSuaChua = dm.NguyenGia * dm.SuaChua / 100m / vm.SoCaNam;
                vm.ChiPhiKhac = dm.NguyenGia * dm.ChiPhiKhac / 100m / vm.SoCaNam;

                // Nhân công VH máy (Đa thành phần)
                vm.ChiPhiNhanCong = 0;
                var tpNC = dm.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = _allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null)
                    {
                        vm.ChiPhiNhanCong += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                    }
                }
                
                vm.HeSoNhienLieuPhu = dm.HeSoNhienLieuPhu;
                vm.NhanCongString = dm.NhanCongString;
            }

            _mayViewModels.Add(vm);
        }
        
        _mayViewModels = _mayViewModels.OrderBy(x => x.MaHieu).ToList();
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

        lblCount = new Label 
        { 
            AutoSize = true, 
            ForeColor = Color.Gray,
            Margin = new Padding(0, 5, 0, 0)
        };

        topPanel.Controls.Add(lblSearch);
        topPanel.Controls.Add(txtSearch);
        topPanel.Controls.Add(lblHint);

        var rightPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(0, 10, 10, 0)
        };

        var btnAutoFit = new Button
        {
            Text = "↕ Tự động giãn cột/dòng",
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(240, 240, 240),
            Font = new Font("Be Vietnam Pro", 9f),
            Margin = new Padding(15, 0, 0, 0)
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

        rightPanel.Controls.Add(btnAutoFit);
        rightPanel.Controls.Add(lblCount);
        topPanel.Controls.Add(rightPanel);

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
        
        // Máy thi công tab with fuel price panel
        var mayContainer = new Panel { Dock = DockStyle.Fill };
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8, 6, 8, 6) };
        fuelPanel.BackColor = Color.FromArgb(240, 248, 255);
        
        int px = 10;
        
        var lblVung = new Label { Text = "Vùng áp dụng:", AutoSize = true, Location = new Point(px, 10), Font = new Font("Be Vietnam Pro", 9f) };
        fuelPanel.Controls.Add(lblVung);
        px += lblVung.PreferredWidth + 4;
        
        var cbVung = new ComboBox { Width = 110, Location = new Point(px, 7), Font = new Font("Be Vietnam Pro", 9f), DropDownStyle = ComboBoxStyle.DropDownList };
        cbVung.Items.AddRange(new string[] { "Vùng II", "Vùng III", "Vùng IV", "Cù Lao Chàm" });
        cbVung.SelectedIndex = 0;
        cbVung.SelectedIndexChanged += (s, e) => {
            _vungApDung = (AIE.Core.Enums.Vung)(cbVung.SelectedIndex + 2);
            RecalculateMachineCosts();
        };
        fuelPanel.Controls.Add(cbVung);
        px += 125;
        
        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px, "");
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px, "");
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px, "");
        
        mayContainer.Controls.Add(dgvMay);
        mayContainer.Controls.Add(fuelPanel);
        tabMay.Controls.Add(mayContainer);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);
        
        UIHelper.ApplyStyle(tabControl);

        tabControl.SelectedIndexChanged += (s, e) => {
            txtSearch.Text = "";
            ApplyFilter();
        };

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

    private TextBox AddFuelInput(Panel panel, string label, ref int x, string defaultValue)
    {
        var lbl = new Label { Text = label, AutoSize = true, Location = new Point(x, 10), Font = new Font("Be Vietnam Pro", 9f) };
        panel.Controls.Add(lbl);
        x += lbl.PreferredWidth + 4;
        var txt = new TextBox { Width = 90, Location = new Point(x, 7), Font = new Font("Be Vietnam Pro", 9.5f), TextAlign = HorizontalAlignment.Right, Text = defaultValue };
        txt.TextChanged += (s, e) => RecalculateMachineCosts();
        panel.Controls.Add(txt);
        x += 100;
        return txt;
    }

    private DataGridView CreateGrid()
    {
        var dgv = new DataGridView
        {
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
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeight = 60,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Regular),
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True },
            EditMode = DataGridViewEditMode.EditOnEnter
        };

        dgv.CellPainting += Dgv_CellPainting;

        UIHelper.ApplyStyle(dgv);

        dgv.CellFormatting += Dgv_CellFormatting;
        
        return dgv;
    }

    private void Dgv_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        if (sender == dgvMay)
        {
            AIE.ExcelAddIn.Helpers.GridHelper.PaintMergedHeader(sender, e, dgvMay, 3, 5, "Định mức", 10, 12, "Chi phí");
        }
    }

    private void LoadData()
    {
        // === VL ===
        dgvVL.Columns.Clear();
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaVL", HeaderText = "Mã VL", DataPropertyName = "MaVL", ReadOnly = true, FillWeight = 15 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenVL", HeaderText = "Tên vật liệu", DataPropertyName = "TenVL", ReadOnly = true, FillWeight = 40 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 10, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGia", HeaderText = "Đơn giá", DataPropertyName = "DonGia", FillWeight = 18, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", DataPropertyName = "GhiChu", FillWeight = 17 });

        // === NC ===
        dgvNC.Columns.Clear();
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaNC", HeaderText = "Mã NC", DataPropertyName = "MaNC", ReadOnly = true, FillWeight = 15 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenNC", HeaderText = "Tên nhân công", DataPropertyName = "TenNC", ReadOnly = false, FillWeight = 30 });
        
        var loaiNCList = new List<KeyValuePair<AIE.Core.Enums.LoaiNhanCong, string>>
        {
            new KeyValuePair<AIE.Core.Enums.LoaiNhanCong, string>(AIE.Core.Enums.LoaiNhanCong.XayDung, "Xây dựng"),
            new KeyValuePair<AIE.Core.Enums.LoaiNhanCong, string>(AIE.Core.Enums.LoaiNhanCong.VanHanhMay, "Vận hành máy"),
            new KeyValuePair<AIE.Core.Enums.LoaiNhanCong, string>(AIE.Core.Enums.LoaiNhanCong.Khac, "Khác")
        };

        var colLoaiNC = new DataGridViewComboBoxColumn 
        { 
            Name = "LoaiNhanCong", 
            HeaderText = "Loại NC", 
            DataPropertyName = "LoaiNhanCong", 
            DataSource = loaiNCList, 
            DisplayMember = "Value",
            ValueMember = "Key",
            ValueType = typeof(AIE.Core.Enums.LoaiNhanCong), 
            FillWeight = 20 
        };
        dgvNC.Columns.Add(colLoaiNC);
        
        // Nhóm column as ComboBox (cascading based on LoaiNhanCong)
        var nhomList = new List<KeyValuePair<int, string>>();
        for (int i = 1; i <= 6; i++) nhomList.Add(new KeyValuePair<int, string>(i, $"Nhóm {i}"));
        var colNhom = new DataGridViewComboBoxColumn 
        { 
            Name = "Nhom", 
            HeaderText = "Nhóm", 
            DataPropertyName = "Nhom", 
            DataSource = nhomList,
            DisplayMember = "Value",
            ValueMember = "Key",
            ValueType = typeof(int),
            FillWeight = 10,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        };
        dgvNC.Columns.Add(colNhom);

        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, FillWeight = 8, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGiaVung2", HeaderText = "Giá Vùng II", DataPropertyName = "DonGiaVung2", FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGiaVung3", HeaderText = "Giá Vùng III", DataPropertyName = "DonGiaVung3", FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGiaVung4", HeaderText = "Giá Vùng IV", DataPropertyName = "DonGiaVung4", FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGiaCLC", HeaderText = "Giá CLC", DataPropertyName = "DonGiaCLC", FillWeight = 12, DefaultCellStyle = new DataGridViewCellStyle { Format = "#,##0", Alignment = DataGridViewContentAlignment.MiddleRight } });

        // Nhóm cascading events
        dgvNC.CurrentCellDirtyStateChanged += DgvNC_CurrentCellDirtyStateChanged;
        dgvNC.CellBeginEdit += DgvNC_CellBeginEdit;
        dgvNC.DataError += (s, e) => { e.ThrowException = false; };

        // === May ===
        dgvMay.Columns.Clear();
        dgvMay.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã hiệu", DataPropertyName = "MaHieu", ReadOnly = true, Width = 90, Frozen = true });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên máy và thiết bị", DataPropertyName = "Ten", ReadOnly = true, Width = 200, Frozen = true, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoCaNam", HeaderText = "Số ca/năm", DataPropertyName = "SoCaNam", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHao", HeaderText = "ĐM Khấu hao", DataPropertyName = "TyLeKhauHao", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChua", HeaderText = "ĐM Sửa chữa", DataPropertyName = "TyLeSuaChua", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhacDM", HeaderText = "ĐM Khác (%)", DataPropertyName = "TyLeKhac", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucNhienLieuDisplay", HeaderText = "Định mức tiêu hao NL", DataPropertyName = "DinhMucNhienLieuDisplay", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, WrapMode = DataGridViewTriState.True } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "HeSoNhienLieuPhu", HeaderText = "Hệ số NL phụ", DataPropertyName = "HeSoNhienLieuPhu", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##", Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhanCongVanHanhDisplay", HeaderText = "Nhân công vận hành", DataPropertyName = "NhanCongVanHanhDisplay", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = false, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhauHao", HeaderText = "CP Khấu hao", DataPropertyName = "ChiPhiKhauHao", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiSuaChua", HeaderText = "CP Sửa chữa", DataPropertyName = "ChiPhiSuaChua", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhacThuc", HeaderText = "CP Khác", DataPropertyName = "ChiPhiKhac", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhienLieu", HeaderText = "CP Nhiên liệu", DataPropertyName = "ChiPhiNhienLieu", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiNhanCong", HeaderText = "Lương thợ", DataPropertyName = "ChiPhiNhanCong", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonGiaCaMay", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "DonGiaCaMay", ReadOnly = true, Width = 130, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });
        
        dgvMay.CellValueChanged += DgvMay_CellValueChanged;
        dgvMay.CellDoubleClick += DgvMay_CellDoubleClick;
        ApplyFilter();
    }

    private List<KeyValuePair<int, string>> GetNhomListForLoai(AIE.Core.Enums.LoaiNhanCong loai)
    {
        var result = new List<KeyValuePair<int, string>>();
        switch (loai)
        {
            case AIE.Core.Enums.LoaiNhanCong.XayDung:
                for (int i = 1; i <= 4; i++) result.Add(new KeyValuePair<int, string>(i, $"Nhóm {i}"));
                break;
            case AIE.Core.Enums.LoaiNhanCong.VanHanhMay:
                for (int i = 1; i <= 6; i++) result.Add(new KeyValuePair<int, string>(i, $"Nhóm {i}"));
                break;
            case AIE.Core.Enums.LoaiNhanCong.Khac:
                for (int i = 1; i <= 3; i++) result.Add(new KeyValuePair<int, string>(i, $"Nhóm {i}"));
                break;
        }
        return result;
    }

    private void DgvNC_CurrentCellDirtyStateChanged(object sender, EventArgs e)
    {
        if (dgvNC.IsCurrentCellDirty && dgvNC.CurrentCell.OwningColumn.Name == "LoaiNhanCong")
        {
            dgvNC.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void DgvNC_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
    {
        if (dgvNC.Columns[e.ColumnIndex].Name == "Nhom")
        {
            var cell = (DataGridViewComboBoxCell)dgvNC[e.ColumnIndex, e.RowIndex];
            var loaiVal = dgvNC["LoaiNhanCong", e.RowIndex].Value;
            
            AIE.Core.Enums.LoaiNhanCong loai = AIE.Core.Enums.LoaiNhanCong.XayDung;
            if (loaiVal != null && Enum.TryParse(loaiVal.ToString(), out AIE.Core.Enums.LoaiNhanCong parsed))
                loai = parsed;

            cell.DataSource = GetNhomListForLoai(loai);
        }
    }

    private void RecalculateMachineCosts()
    {
        if (_suppressRecalc) return;

        decimal.TryParse(txtGiaXang?.Text, out decimal gx);
        decimal.TryParse(txtGiaDiezel?.Text, out decimal gdz);
        decimal.TryParse(txtGiaDien?.Text, out decimal gdi);
        
        _suppressRecalc = true;
        foreach (var m in _mayViewModels)
        {
            // Tính chi phí khấu hao, sửa chữa, khác
            if (m.SoCaNam > 0 && m.NguyenGia > 0)
            {
                m.ChiPhiKhauHao = m.NguyenGia * m.TyLeKhauHao / 100m / m.SoCaNam;
                m.ChiPhiSuaChua = m.NguyenGia * m.TyLeSuaChua / 100m / m.SoCaNam;
                m.ChiPhiKhac = m.NguyenGia * m.TyLeKhac / 100m / m.SoCaNam;
            }

            // Tính chi phí nhiên liệu
            m.ChiPhiNhienLieu = (m.DinhMucXang * gx + m.DinhMucDiezel * gdz + m.DinhMucDien * gdi) * m.HeSoNhienLieuPhu;
            
            // Tính lương thợ (Nhân công vận hành đa thành phần)
            m.ChiPhiNhanCong = 0;
            var dm = _mayDmRepo.GetByMaMay(m.MaHieu);
            if (dm != null)
            {
                var tpNC = dm.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = _allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null)
                    {
                        m.ChiPhiNhanCong += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                    }
                }
            }
        }
        dgvMay.Refresh();
        _suppressRecalc = false;
    }

    private decimal ParseVn(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        if (decimal.TryParse(text, System.Globalization.NumberStyles.Any, new System.Globalization.CultureInfo("vi-VN"), out var val)) return val;
        return 0;
    }

    private void DgvMay_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        
        var colName = dgvMay.Columns[e.ColumnIndex].Name;
        if (colName == "KhauHao" || colName == "SuaChua" || colName == "ChiPhiKhacDM" || colName == "NguyenGia")
        {
            var vm = (DonGiaMayViewModel)dgvMay.Rows[e.RowIndex].DataBoundItem;
            var dm = _mayDmRepo.GetByMaMay(vm.MaHieu);
            if (dm != null)
            {
                dm.NguyenGia = ParseVn(vm.NguyenGia.ToString());
                dm.KhauHao = ParseVn(vm.TyLeKhauHao.ToString());
                dm.SuaChua = ParseVn(vm.TyLeSuaChua.ToString());
                dm.ChiPhiKhac = ParseVn(vm.TyLeKhac.ToString());
                
                _mayDmRepo.Upsert(dm);
                RecalculateMachineCosts();
                
                vm.ChiPhiKhauHao = dm.NguyenGia * dm.KhauHao / 100m / vm.SoCaNam;
                vm.ChiPhiSuaChua = dm.NguyenGia * dm.SuaChua / 100m / vm.SoCaNam;
                vm.ChiPhiKhac = dm.NguyenGia * dm.ChiPhiKhac / 100m / vm.SoCaNam;
                
                dgvMay.InvalidateRow(e.RowIndex);
            }
        }
    }

    private void TxtSearch_TextChanged(object sender, EventArgs e)
    {
        ApplyFilter();
    }

    private string _lastFilterVL = null;
    private string _lastFilterNC = null;
    private string _lastFilterMay = null;

    private void ApplyFilter()
    {
        string keyword = txtSearch.Text.Trim().ToLower();

        if (tabControl.SelectedIndex == 0 && keyword != _lastFilterVL)
        {
            var filteredVL = string.IsNullOrEmpty(keyword)
                ? _allVL
                : _allVL.Where(x => MatchAllKeywords(keyword, x.MaVL, x.TenVL)).ToList();
            dgvVL.DataSource = new BindingSource { DataSource = filteredVL };
            dgvVL.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            lblCount.Text = $"Hiển thị {filteredVL.Count} / {_allVL.Count} vật liệu";
            _lastFilterVL = keyword;
        }
        else if (tabControl.SelectedIndex == 1 && keyword != _lastFilterNC)
        {
            IEnumerable<AIE.Core.Models.NhanCong> filteredNC = _allNC;
            
            if (!string.IsNullOrEmpty(keyword))
            {
                filteredNC = filteredNC.Where(x => MatchAllKeywords(keyword, x.MaNC, x.TenNC));
            }
            
            var sortedNC = filteredNC.OrderBy(x => x.LoaiNhanCong).ThenBy(x => x.Nhom).ToList();
            dgvNC.DataSource = new BindingSource { DataSource = sortedNC };
            dgvNC.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            lblCount.Text = $"Hiển thị {sortedNC.Count} / {_allNC.Count} nhân công";
            _lastFilterNC = keyword;
        }
        else if (tabControl.SelectedIndex == 2 && keyword != _lastFilterMay)
        {
            var filteredMay = string.IsNullOrEmpty(keyword)
                ? _mayViewModels
                : _mayViewModels.Where(x => MatchAllKeywords(keyword, x.MaHieu, x.Ten)).ToList();
            dgvMay.DataSource = new BindingSource { DataSource = filteredMay };
            dgvMay.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            lblCount.Text = $"Hiển thị {filteredMay.Count} / {_mayViewModels.Count} máy thi công";
            _lastFilterMay = keyword;
        }
    }

    private static bool MatchAllKeywords(string keyword, params string[] fields)
    {
        string combined = string.Join(" ", fields).ToLower();
        var words = keyword.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return words.All(w => combined.Contains(w));
    }

    private void DgvMay_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var maMay = dgvMay["MaHieu", e.RowIndex].Value?.ToString();
        var tenMay = dgvMay["Ten", e.RowIndex].Value?.ToString();
        if (!string.IsNullOrEmpty(maMay))
        {
            using var frm = new NhapDinhMucCaMayForm(maMay, tenMay ?? "", _mayDmRepo);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                var dmMay = _mayDmRepo.GetByMaMay(maMay);
                var mayM = _mayViewModels.FirstOrDefault(x => x.MaHieu == maMay);
                if (dmMay != null && mayM != null)
                {
                    mayM.SoCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                    mayM.NguyenGia = dmMay.NguyenGia;
                    mayM.TyLeKhauHao = dmMay.KhauHao;
                    mayM.TyLeSuaChua = dmMay.SuaChua;
                    mayM.TyLeKhac = dmMay.ChiPhiKhac;
                    mayM.NhomNhanCong = dmMay.NhomNhanCong;
                    mayM.SoLuongNhanCong = dmMay.SoLuongNhanCong;
                    mayM.HeSoNhienLieuPhu = dmMay.HeSoNhienLieuPhu;
                    mayM.NhanCongString = dmMay.NhanCongString;
                    mayM.DinhMucXang = dmMay.DinhMucXang;
                    mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                    mayM.DinhMucDien = dmMay.DinhMucDien;
                }
                RecalculateMachineCosts();
            }
        }
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

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
        if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal val))
        {
            if (val == 0)
            {
                e.Value = "-";
                e.FormattingApplied = true;
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }
    }
}

/// <summary>
/// ViewModel cho máy thi công hiển thị trong QuanLyDonGiaForm.
/// Kết hợp dữ liệu MayThiCong + DinhMucCaMay_TT37 để tính giá ca máy trực tiếp.
/// </summary>
public class DonGiaMayViewModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    
    // Reference to original DB model (for saving)
    public MayThiCong MayThiCong { get; set; }
    
    // TT37 factors
    public decimal NguyenGia { get; set; }
    public int SoCaNam { get; set; } = 250;
    public decimal TyLeKhauHao { get; set; }
    public decimal TyLeSuaChua { get; set; }
    public decimal TyLeKhac { get; set; }
    public decimal DinhMucXang { get; set; }
    public decimal DinhMucDiezel { get; set; }
    public decimal DinhMucDien { get; set; }
    public decimal SoLuongNhanCong { get; set; }
    public int NhomNhanCong { get; set; }

    // Computed cost components
    public decimal ChiPhiKhauHao { get; set; }
    public decimal ChiPhiSuaChua { get; set; }
    public decimal ChiPhiKhac { get; set; }
    public decimal ChiPhiNhienLieu { get; set; }
    public decimal ChiPhiNhanCong { get; set; }
    
    // Total
    public decimal GiaHienTruong => ChiPhiKhauHao + ChiPhiSuaChua + ChiPhiKhac + ChiPhiNhienLieu + ChiPhiNhanCong;
    public decimal DonGiaCaMay => GiaHienTruong;

    // Display strings
    public string DinhMucNhienLieuDisplay
    {
        get
        {
            var parts = new List<string>();
            if (DinhMucXang > 0) parts.Add($"{DinhMucXang:#.##} lít xăng");
            if (DinhMucDiezel > 0) parts.Add($"{DinhMucDiezel:#.##} lít diezel");
            if (DinhMucDien > 0) parts.Add($"{DinhMucDien:#.##} kWh");
            return string.Join(" + ", parts);
        }
    }
    
    public decimal HeSoNhienLieuPhu { get; set; } = 1.0m;
    public string NhanCongString { get; set; }

    public string NhanCongVanHanhDisplay
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(NhanCongString))
                return NhanCongString;

            if (SoLuongNhanCong <= 0) return "";
            string tenTho = NhomNhanCong switch
            {
                1 => "nhân công vận hành",
                2 => "lái xe",
                3 => "thủy thủ, thợ máy, thợ điện",
                4 => "máy trưởng, thuyền trưởng",
                5 => "máy trưởng tàu biển",
                6 => "thuyền trưởng, thuyền phó",
                _ => $"nhân công nhóm {NhomNhanCong}"
            };
            return $"{SoLuongNhanCong:#.##} {tenTho}";
        }
    }
}
