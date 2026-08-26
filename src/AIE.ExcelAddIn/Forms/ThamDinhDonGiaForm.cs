using AIE.Core.Models;
using AIE.Data;
using AIE.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms;

public class ThamDinhDonGiaForm : Form
{
    private TabControl tabControl;
    private DataGridView dgvVL, dgvNC, dgvMay;
    private TextBox txtTenBoDonGia;
    private TextBox txtGiaXang, txtGiaDiezel, txtGiaDien;
    
    private List<DgVatLieuModel> _vatLieuList;
    private List<DgNhanCongModel> _nhanCongList;
    private List<DgMayThiCongModel> _mayThiCongList;
    
    private readonly CultureInfo ViVn = new CultureInfo("vi-VN");
    private bool _suppressRecalc = false;

    private AIE.Core.Enums.Vung _vungApDung;

    public ThamDinhDonGiaForm(List<VatTuGiaModel> extractedItems, AIE.Core.Enums.Vung vung = AIE.Core.Enums.Vung.VungII)
    {
        _vungApDung = vung;
        InitializeComponent();
        PrepareData(extractedItems);
        LoadDataToGrids();
    }

    private void InitializeComponent()
    {
        this.Text = "Bộ Đơn Giá Thẩm Định";
        var workingArea = Screen.PrimaryScreen.WorkingArea;
        this.Size = new Size((int)(workingArea.Width * 0.8), (int)(workingArea.Height * 0.8));
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Font = new Font("Be Vietnam Pro", 9.5f);

        // Header Panel (Tên bộ đơn giá & Nút lưu)
        var topPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(12) };
        var lblTitle = new Label { Text = "Tên Bộ Đơn Giá:", AutoSize = true, Location = new Point(12, 18), Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        txtTenBoDonGia = new TextBox { Location = new Point(140, 15), Width = 400, Text = $"Đơn giá thẩm định - {DateTime.Now:dd/MM/yyyy HH:mm}" };
        
        var btnLuu = new Button { Text = "Lưu Đơn Giá", Size = new Size(120, 32), Location = new Point(560, 13), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };
        btnLuu.FlatAppearance.BorderSize = 0;
        btnLuu.Click += BtnLuu_Click;

        topPanel.Controls.Add(lblTitle);
        topPanel.Controls.Add(txtTenBoDonGia);
        topPanel.Controls.Add(btnLuu);

        tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold) };

        // Tab Vật Liệu
        var tabVL = new TabPage("Vật liệu");
        dgvVL = CreateGrid();
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã VL", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên vật liệu", DataPropertyName = "Ten", ReadOnly = true, Width = 300 });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaGoc", HeaderText = "Giá gốc", DataPropertyName = "GiaGoc", Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "CuocVC", HeaderText = "Cước VC", DataPropertyName = "CuocVC", Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow } });
        dgvVL.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá hiện trường", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvVL.Font, FontStyle.Bold) } });
        dgvVL.CellValueChanged += DgvVL_CellValueChanged;
        tabVL.Controls.Add(dgvVL);

        // Tab Nhân Công
        var tabNC = new TabPage("Nhân công");
        dgvNC = CreateGrid();
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã NC", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên nhân công", DataPropertyName = "Ten", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });

        var nhomList = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>
        {
            new System.Collections.Generic.KeyValuePair<int, string>(1, "Nhóm I"),
            new System.Collections.Generic.KeyValuePair<int, string>(2, "Nhóm II"),
            new System.Collections.Generic.KeyValuePair<int, string>(3, "Nhóm III"),
            new System.Collections.Generic.KeyValuePair<int, string>(4, "Nhóm IV")
        };
        var colNhom = new DataGridViewComboBoxColumn 
        { 
            Name = "NhomNhanCong", 
            HeaderText = "Nhóm nhân công", 
            DataPropertyName = "NhomNhanCong", 
            DataSource = nhomList, 
            DisplayMember = "Value",
            ValueMember = "Key",
            ValueType = typeof(int), 
            Width = 140 
        };
        dgvNC.Columns.Add(colNhom);

        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá nhân công (Editable)", DataPropertyName = "GiaHienTruong", Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow, ForeColor = Color.Red, Font = new Font(dgvNC.Font, FontStyle.Bold) } });
        
        var colBtn = new DataGridViewButtonColumn
        {
            Name = "LuuGiaGoc",
            HeaderText = "",
            Text = "Cập nhật vào giá gốc",
            UseColumnTextForButtonValue = true,
            Width = 150
        };
        dgvNC.Columns.Add(colBtn);

        dgvNC.CurrentCellDirtyStateChanged += DgvNC_CurrentCellDirtyStateChanged;
        dgvNC.CellValueChanged += DgvNC_CellValueChanged;
        dgvNC.CellContentClick += DgvNC_CellContentClick;

        var lblNCWarning = new Label { Text = "Ghi chú: Nếu muốn cập nhật giá gốc cho toàn hệ thống, vui lòng đổi tại mục Quản lý Đơn giá hoặc dùng nút bên dưới. Sửa trực tiếp chỉ áp dụng cho bảng này.", Dock = DockStyle.Top, Height = 30, ForeColor = Color.DimGray, Font = new Font("Be Vietnam Pro", 9f, FontStyle.Italic), TextAlign = ContentAlignment.MiddleLeft };
        tabNC.Controls.Add(dgvNC);
        tabNC.Controls.Add(lblNCWarning);

        // Tab Máy thi công
        var tabMay = new TabPage("Máy thi công");
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(10), BackColor = Color.FromArgb(240, 248, 255) };
        int px = 10;
        
        // Vùng áp dụng dropdown
        var lblVung = new Label { Text = "Vùng áp dụng:", AutoSize = true, Location = new Point(px, 20), Font = new Font("Be Vietnam Pro", 9f) };
        fuelPanel.Controls.Add(lblVung);
        px += lblVung.PreferredWidth + 4;
        
        var cbVungMay = new ComboBox { Width = 110, Location = new Point(px, 17), Font = new Font("Be Vietnam Pro", 9f), DropDownStyle = ComboBoxStyle.DropDownList };
        cbVungMay.Items.AddRange(new string[] { "Vùng II", "Vùng III", "Vùng IV", "Cù Lao Chàm" });
        cbVungMay.SelectedIndex = (int)_vungApDung - 2;
        cbVungMay.SelectedIndexChanged += (s, e) => {
            _vungApDung = (AIE.Core.Enums.Vung)(cbVungMay.SelectedIndex + 2);
            RecalculateMachineCosts();
        };
        fuelPanel.Controls.Add(cbVungMay);
        px += 125;
        
        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px, "20000");
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px, "18000");
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px, "2000");

        dgvMay = CreateGrid();
        dgvMay.ColumnHeadersHeight = 45;
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã Máy", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên máy", DataPropertyName = "Ten", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoCaNam", HeaderText = "Số ca/năm", DataPropertyName = "SoCaNam", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHao", HeaderText = "ĐM Khấu hao", DataPropertyName = "TyLeKhauHao", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChua", HeaderText = "ĐM Sửa chữa", DataPropertyName = "TyLeSuaChua", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhac", HeaderText = "ĐM Khác (%)", DataPropertyName = "TyLeKhac", ReadOnly = false, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, BackColor = Color.FromArgb(255, 255, 200) } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DinhMucNhienLieuDisplay", HeaderText = "Định mức tiêu hao NL", DataPropertyName = "DinhMucNhienLieuDisplay", ReadOnly = true, Width = 150, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter, WrapMode = DataGridViewTriState.True } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "HeSoNhienLieuPhu", HeaderText = "Hệ số NL phụ", DataPropertyName = "HeSoNhienLieuPhu", ReadOnly = true, Width = 70, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.##", Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhanCongVanHanhDisplay", HeaderText = "Nhân công vận hành", DataPropertyName = "NhanCongVanHanhDisplay", ReadOnly = true, Width = 250, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft, WrapMode = DataGridViewTriState.True } });
        
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = false, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.FromArgb(255, 255, 200) } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHaoGia", HeaderText = "CP Khấu hao", DataPropertyName = "KhauHao", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChuaGia", HeaderText = "CP Sửa chữa", DataPropertyName = "SuaChua", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhacGia", HeaderText = "CP Khác", DataPropertyName = "ChiPhiKhacGia", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhienLieu", HeaderText = "CP Nhiên liệu", DataPropertyName = "NhienLieu", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "LuongTho", HeaderText = "Lương thợ", DataPropertyName = "LuongTho", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 120, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });
        
        dgvMay.CellPainting += DgvMay_CellPainting;
        dgvMay.CellDoubleClick += DgvMay_CellDoubleClick;
        dgvMay.CellValueChanged += DgvMay_CellValueChanged;
        var ctxMenu = new ContextMenuStrip();
        tabMay.Controls.Add(dgvMay);
        tabMay.Controls.Add(fuelPanel);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);

        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
    }

    private TextBox AddFuelInput(Control parent, string labelText, ref int x, string defaultValue)
    {
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(x, 20), Font = new Font("Be Vietnam Pro", 9f) };
        int lblWidth = lbl.PreferredWidth;
        var txt = new TextBox { Text = defaultValue, Width = 80, Location = new Point(x + lblWidth + 5, 17), Font = new Font("Be Vietnam Pro", 9f), TextAlign = HorizontalAlignment.Right };
        txt.TextChanged += (s, e) => RecalculateMachineCosts();
        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        x += lblWidth + 5 + txt.Width + 20;
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
            BackgroundColor = Color.White,
            RowTemplate = { Height = 28 },
            Font = new Font("Be Vietnam Pro", 9.5f),
            EnableHeadersVisualStyles = false,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
            RowHeadersVisible = true,
            RowHeadersWidth = 25,
            AllowUserToResizeRows = true,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        };
        
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 235, 235);
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro", 9f, FontStyle.Bold);
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.ColumnHeadersHeight = 35;
        
        return dgv;
    }

    private void PrepareData(List<VatTuGiaModel> extractedItems)
    {
        var db = new DatabaseManager();
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);
        var mayRepo = new MayThiCongRepository(db.Context);
        var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);

        _vatLieuList = new List<DgVatLieuModel>();
        _nhanCongList = new List<DgNhanCongModel>();
        _mayThiCongList = new List<DgMayThiCongModel>();

        foreach (var item in extractedItems)
        {
            if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.VL)
            {
                _vatLieuList.Add(new DgVatLieuModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi,
                    GiaGoc = 0, // Theo yêu cầu: vật liệu không lấy giá gốc từ danh mục chung
                    CuocVC = 0
                });
            }
            else if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.NC)
            {
                var ncDb = ncRepo.GetAll().FirstOrDefault(x => x.MaNC == item.MaHieu);
                _nhanCongList.Add(new DgNhanCongModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi,
                    GiaHienTruong = item.GiaChuan ?? 0,
                    NhomNhanCong = ncDb != null ? ncDb.Nhom : 1
                });
            }
            else if (item.LoaiHP == AIE.Core.Enums.LoaiHaoPhi.MAY)
            {
                var dmMay = dmMayRepo.GetByMaMay(item.MaHieu);
                var mayM = new DgMayThiCongModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi
                };
                
                if (dmMay != null)
                {
                    // Tạm tính giá trị theo TT37
                    mayM.SoCaNam = dmMay.SoCaNam > 0 ? dmMay.SoCaNam : 250;
                    mayM.NguyenGia = dmMay.NguyenGia;
                    mayM.TyLeKhauHao = dmMay.KhauHao;
                    mayM.TyLeSuaChua = dmMay.SuaChua;
                    mayM.TyLeKhac = dmMay.ChiPhiKhac;
                    mayM.NhomNhanCong = dmMay.NhomNhanCong;
                    mayM.SoLuongNhanCong = dmMay.SoLuongNhanCong;
                    mayM.HeSoNhienLieuPhu = dmMay.HeSoNhienLieuPhu;
                    mayM.NhanCongString = dmMay.NhanCongString;
                    
                    decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                    mayM.KhauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / mayM.SoCaNam;
                    
                    mayM.SuaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / mayM.SoCaNam;
                    mayM.ChiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / mayM.SoCaNam;
                    
                    mayM.DinhMucXang = dmMay.DinhMucXang;
                    mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                    mayM.DinhMucDien = dmMay.DinhMucDien;
                    
                    mayM.LuongTho = 0; // Thợ lái (Đa thành phần)
                    var allNC = ncRepo.GetAll();
                    var tpNC = dmMay.GetThanhPhanNhanCong();
                    foreach (var tp in tpNC)
                    {
                        var ncMay = allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                        if (ncMay != null) mayM.LuongTho += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                    }
                }
                else
                {
                    // Lấy default từ bảng MayThiCong
                    var mg = mayRepo.GetByMa(item.MaHieu);
                    if (mg != null)
                    {
                        mayM.KhauHao = mg.DonGia; // Lấy tạm giá gốc đẩy vào 1 cục
                    }
                }

                _mayThiCongList.Add(mayM);
            }
        }
    }

    private void LoadDataToGrids()
    {
        dgvVL.DataSource = new BindingSource { DataSource = _vatLieuList };
        dgvNC.DataSource = new BindingSource { DataSource = _nhanCongList };
        dgvMay.DataSource = new BindingSource { DataSource = _mayThiCongList };
        
        dgvVL.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        dgvNC.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        dgvMay.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        RecalculateMachineCosts();
    }

    private void DgvNC_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (dgvNC.IsCurrentCellDirty && dgvNC.CurrentCell is DataGridViewComboBoxCell)
        {
            dgvNC.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void DgvNC_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        
        if (dgvNC.Columns[e.ColumnIndex].Name == "NhomNhanCong")
        {
            var nhom = dgvNC["NhomNhanCong", e.RowIndex].Value;
            if (nhom is int n)
            {
                var db = new AIE.Data.DatabaseManager();
                var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
                var allNC = ncRepo.GetAll();
                var ncDb = allNC.FirstOrDefault(x => x.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.XayDung && x.Nhom == n);
                if (ncDb != null)
                {
                    _nhanCongList[e.RowIndex].GiaHienTruong = ncDb.GetDonGia(_vungApDung);
                    dgvNC.InvalidateRow(e.RowIndex);
                }
            }
        }
    }

    private void DgvNC_CellContentClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        
        if (dgvNC.Columns[e.ColumnIndex].Name == "LuuGiaGoc")
        {
            var maNC = _nhanCongList[e.RowIndex].MaHieu;
            var currentPrice = _nhanCongList[e.RowIndex].GiaHienTruong;
            
            var db = new AIE.Data.DatabaseManager();
            var ncRepo = new AIE.Data.Repositories.NhanCongRepository(db.Context);
            var ncDb = ncRepo.GetAll().FirstOrDefault(x => x.MaNC == maNC);
            if (ncDb != null)
            {
                ncDb.SetDonGia(_vungApDung, currentPrice);
                ncDb.NgayCapNhat = DateTime.Now;
                ncRepo.Upsert(ncDb);
                MessageBox.Show($"Đã cập nhật giá gốc cho nhân công {maNC} = {currentPrice:N0}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Không tìm thấy mã nhân công trong cơ sở dữ liệu gốc.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void DgvVL_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_suppressRecalc || e.RowIndex < 0) return;
        dgvVL.InvalidateRow(e.RowIndex);
    }

    private void RecalculateMachineCosts()
    {
        if (_suppressRecalc) return;
        
        decimal.TryParse(txtGiaXang.Text, out decimal gx);
        decimal.TryParse(txtGiaDiezel.Text, out decimal gdz);
        decimal.TryParse(txtGiaDien.Text, out decimal gdi);
        
        _suppressRecalc = true;
        
        var db = new DatabaseManager();
        var ncRepo = new NhanCongRepository(db.Context);
        var dmMayRepo = new DinhMucCaMayRepository(db.Context);
        var allNC = ncRepo.GetAll();
        
        foreach (var m in _mayThiCongList)
        {
            m.NhienLieu = (m.DinhMucXang * gx + m.DinhMucDiezel * gdz + m.DinhMucDien * gdi) * m.HeSoNhienLieuPhu;
            
            var dmMay = dmMayRepo.GetByMaMay(m.MaHieu);
            if (dmMay != null)
            {
                m.LuongTho = 0;
                var tpNC = dmMay.GetThanhPhanNhanCong();
                foreach (var tp in tpNC)
                {
                    var ncMay = allNC.FirstOrDefault(n => n.LoaiNhanCong == AIE.Core.Enums.LoaiNhanCong.VanHanhMay && n.Nhom == tp.Nhom);
                    if (ncMay != null) m.LuongTho += ncMay.GetDonGia(_vungApDung) * tp.SoLuong;
                }
            }
        }
        dgvMay.Refresh();
        _suppressRecalc = false;
    }

    private void DgvMay_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        
        var colName = dgvMay.Columns[e.ColumnIndex].Name;
        if (colName == "KhauHao" || colName == "SuaChua" || colName == "ChiPhiKhac" || colName == "NguyenGia")
        {
            var m = _mayThiCongList[e.RowIndex];
            var db = new DatabaseManager();
            var dmMayRepo = new DinhMucCaMayRepository(db.Context);
            var dmMay = dmMayRepo.GetByMaMay(m.MaHieu);
            
            if (dmMay != null)
            {
                dmMay.NguyenGia = m.NguyenGia;
                dmMay.KhauHao = m.TyLeKhauHao;
                dmMay.SuaChua = m.TyLeSuaChua;
                dmMay.ChiPhiKhac = (decimal)m.TyLeKhac;
                dmMay.NguyenGia = m.NguyenGia;
                dmMayRepo.Upsert(dmMay);
                RecalculateMachineCosts();
            }
            decimal g_th = m.NguyenGia >= 30000000m ? m.NguyenGia * 0.1m : 0m;
            m.KhauHao = ((m.NguyenGia - g_th) * m.TyLeKhauHao / 100m) / m.SoCaNam;
            m.SuaChua = (m.NguyenGia * m.TyLeSuaChua / 100m) / m.SoCaNam;
            m.ChiPhiKhac = (m.NguyenGia * m.TyLeKhac / 100m) / m.SoCaNam;
            
            dgvMay.InvalidateRow(e.RowIndex);
        }
    }

    private void BtnLuu_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtTenBoDonGia.Text))
        {
            MessageBox.Show("Vui lòng nhập tên Bộ đơn giá", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            dgvVL.EndEdit();
            dgvNC.EndEdit();
            dgvMay.EndEdit();

            var db = new DatabaseManager();
            var repo = new BoDonGiaRepository(db.Context);
            
            decimal.TryParse(txtGiaXang.Text, out decimal gx);
            decimal.TryParse(txtGiaDiezel.Text, out decimal gdz);
            decimal.TryParse(txtGiaDien.Text, out decimal gdi);

            // 1. Create BoDonGia
            int id = repo.Create(txtTenBoDonGia.Text, gx, gdz, gdi, "Thẩm định từ Excel");

            // 2. Save VL
            foreach (var vl in _vatLieuList)
                repo.SaveGiaVL(id, vl.MaHieu, vl.GiaGoc, vl.CuocVC, vl.GiaHienTruong);

            // 3. Save NC
            foreach (var nc in _nhanCongList)
                repo.SaveGiaNC(id, nc.MaHieu, nc.GiaHienTruong);

            // 4. Save May
            foreach (var m in _mayThiCongList)
                repo.SaveGiaMay(id, m.MaHieu, m.GiaHienTruong);

            MessageBox.Show("Đã lưu Bộ Đơn Giá thành công!\nBây giờ bạn có thể chọn Bộ đơn giá này khi sử dụng chức năng Kiểm tra.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show("Lỗi lưu đơn giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    private void DgvMay_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        // Định mức: cols 3, 4, 5  |  Chi phí: cols 10, 11, 12
        AIE.ExcelAddIn.Helpers.GridHelper.PaintMergedHeader(sender, e, dgvMay, 3, 5, "Định mức", 10, 12, "Chi phí");
    }

    private void DgvMay_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var maMay = dgvMay["MaHieu", e.RowIndex].Value?.ToString();
        var tenMay = dgvMay["Ten", e.RowIndex].Value?.ToString();
        if (!string.IsNullOrEmpty(maMay))
        {
            var db = new AIE.Data.DatabaseManager();
            var mayDmRepo = new AIE.Data.Repositories.DinhMucCaMayRepository(db.Context);
            using var frm = new NhapDinhMucCaMayForm(maMay, tenMay ?? "", mayDmRepo);
            if (frm.ShowDialog() == DialogResult.OK)
            {
                RecalculateMachineCosts();
            }
        }
    }
}

public class DgVatLieuModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal GiaGoc { get; set; }
    public decimal CuocVC { get; set; }
    public decimal GiaHienTruong => GiaGoc + CuocVC;
}

public class DgNhanCongModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public int NhomNhanCong { get; set; }
    public decimal GiaHienTruong { get; set; }
}

public class DgMayThiCongModel
{
    public string MaHieu { get; set; } = string.Empty;
    public string Ten { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal NguyenGia { get; set; }
    public decimal KhauHao { get; set; }
    public decimal SuaChua { get; set; }
    public decimal ChiPhiKhac { get; set; }
    public decimal LuongTho { get; set; }
    public decimal NhienLieu { get; set; }
    public decimal GiaHienTruong => KhauHao + SuaChua + ChiPhiKhac + LuongTho + NhienLieu;
    
    // Internal TT37 factors
    public decimal SoCaNam { get; set; } = 250;
    public decimal TyLeKhauHao { get; set; }
    public decimal TyLeSuaChua { get; set; }
    public decimal TyLeKhac { get; set; }
    public decimal DinhMucXang { get; set; }
    public decimal DinhMucDiezel { get; set; }
    public decimal DinhMucDien { get; set; }
    public int NhomNhanCong { get; set; }
    public decimal SoLuongNhanCong { get; set; }
    public decimal HeSoNhienLieuPhu { get; set; } = 1.0m;
    public string NhanCongString { get; set; }
    
    public string DinhMucNhienLieuDisplay
    {
        get
        {
            var parts = new System.Collections.Generic.List<string>();
            if (DinhMucXang > 0) parts.Add($"{DinhMucXang:#.##} lít xăng");
            if (DinhMucDiezel > 0) parts.Add($"{DinhMucDiezel:#.##} lít diezel");
            if (DinhMucDien > 0) parts.Add($"{DinhMucDien:#.##} kWh");
            return string.Join(" + ", parts);
        }
    }
    
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
