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

    public ThamDinhDonGiaForm(List<VatTuGiaModel> extractedItems)
    {
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
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên nhân công", DataPropertyName = "Ten", ReadOnly = true, Width = 300 });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvNC.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "Giá nhân công (Editable)", DataPropertyName = "GiaHienTruong", Width = 200, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, BackColor = Color.LightYellow, ForeColor = Color.Red, Font = new Font(dgvNC.Font, FontStyle.Bold) } });
        
        var lblNCWarning = new Label { Text = "Ghi chú: Nếu muốn cập nhật giá gốc cho toàn hệ thống, vui lòng đổi tại mục Quản lý Đơn giá. Sửa ở đây chỉ áp dụng cho bộ đơn giá này.", Dock = DockStyle.Top, Height = 30, ForeColor = Color.DimGray, Font = new Font("Be Vietnam Pro", 9f, FontStyle.Italic), TextAlign = ContentAlignment.MiddleLeft };
        tabNC.Controls.Add(dgvNC);
        tabNC.Controls.Add(lblNCWarning);

        // Tab Máy thi công
        var tabMay = new TabPage("Máy thi công");
        var fuelPanel = new Panel { Dock = DockStyle.Top, Height = 60, Padding = new Padding(10), BackColor = Color.FromArgb(240, 248, 255) };
        int px = 10;
        txtGiaXang = AddFuelInput(fuelPanel, "Giá Xăng (đ/lít):", ref px, "20000");
        txtGiaDiezel = AddFuelInput(fuelPanel, "Giá Diezel (đ/lít):", ref px, "18000");
        txtGiaDien = AddFuelInput(fuelPanel, "Giá Điện (đ/kWh):", ref px, "2000");

        dgvMay = CreateGrid();
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", HeaderText = "Mã Máy", DataPropertyName = "MaHieu", ReadOnly = true, Width = 100 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten", HeaderText = "Tên máy", DataPropertyName = "Ten", ReadOnly = true, Width = 250 });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "DonVi", HeaderText = "ĐVT", DataPropertyName = "DonVi", ReadOnly = true, Width = 60, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NguyenGia", HeaderText = "Nguyên giá", DataPropertyName = "NguyenGia", ReadOnly = true, Width = 110, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "KhauHao", HeaderText = "Khấu hao", DataPropertyName = "KhauHao", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "SuaChua", HeaderText = "Sửa chữa", DataPropertyName = "SuaChua", ReadOnly = true, Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChiPhiKhac", HeaderText = "CP Khác", DataPropertyName = "ChiPhiKhac", ReadOnly = true, Width = 80, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "LuongTho", HeaderText = "Lương thợ", DataPropertyName = "LuongTho", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhienLieu", HeaderText = "Nhiên liệu", DataPropertyName = "NhienLieu", ReadOnly = true, Width = 100, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight } });
        dgvMay.Columns.Add(new DataGridViewTextBoxColumn { Name = "GiaHienTruong", HeaderText = "ĐƠN GIÁ CA MÁY", DataPropertyName = "GiaHienTruong", ReadOnly = true, Width = 140, DefaultCellStyle = new DataGridViewCellStyle { Format = "N0", FormatProvider = ViVn, Alignment = DataGridViewContentAlignment.MiddleRight, ForeColor = Color.Red, Font = new Font(dgvMay.Font, FontStyle.Bold) } });
        tabMay.Controls.Add(dgvMay);
        tabMay.Controls.Add(fuelPanel);

        tabControl.TabPages.Add(tabVL);
        tabControl.TabPages.Add(tabNC);
        tabControl.TabPages.Add(tabMay);

        this.Controls.Add(tabControl);
        this.Controls.Add(topPanel);
    }

    private TextBox AddFuelInput(Panel parent, string labelText, ref int x, string defVal)
    {
        var lbl = new Label { Text = labelText, AutoSize = true, Location = new Point(x, 20) };
        var txt = new TextBox { Location = new Point(x + lbl.Width + 5, 17), Width = 100, Text = defVal, TextAlign = HorizontalAlignment.Right };
        txt.TextChanged += (s, e) => RecalculateMachineCosts();
        parent.Controls.Add(lbl);
        parent.Controls.Add(txt);
        x += lbl.Width + txt.Width + 25;
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
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
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
                _nhanCongList.Add(new DgNhanCongModel
                {
                    MaHieu = item.MaHieu,
                    Ten = item.TenVatTu,
                    DonVi = item.DonVi,
                    GiaHienTruong = item.GiaChuan ?? 0
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
                    // Tạm tính giá trị theo TT37 (giả định 1 năm = 250 ca)
                    decimal soCa = 250;
                    mayM.NguyenGia = dmMay.NguyenGia;
                    
                    decimal g_th = dmMay.NguyenGia >= 30000000m ? dmMay.NguyenGia * 0.1m : 0m;
                    mayM.KhauHao = ((dmMay.NguyenGia - g_th) * dmMay.KhauHao / 100m) / soCa;
                    
                    mayM.SuaChua = (dmMay.NguyenGia * dmMay.SuaChua / 100m) / soCa;
                    mayM.ChiPhiKhac = (dmMay.NguyenGia * dmMay.ChiPhiKhac / 100m) / soCa;
                    
                    mayM.DinhMucXang = dmMay.DinhMucXang;
                    mayM.DinhMucDiezel = dmMay.DinhMucDiezel;
                    mayM.DinhMucDien = dmMay.DinhMucDien;
                    
                    mayM.LuongTho = 0; // Thợ lái
                    if (dmMay.NhomNhanCong > 0) {
                        var ncRepo = new NhanCongRepository(db.Context);
                        var ncMay = ncRepo.GetAll().FirstOrDefault(n => n.Nhom == dmMay.NhomNhanCong);
                        if (ncMay != null) mayM.LuongTho = ncMay.DonGia * dmMay.SoLuongNhanCong;
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
        dgvVL.DataSource = new System.ComponentModel.BindingList<DgVatLieuModel>(_vatLieuList);
        dgvNC.DataSource = new System.ComponentModel.BindingList<DgNhanCongModel>(_nhanCongList);
        dgvMay.DataSource = new System.ComponentModel.BindingList<DgMayThiCongModel>(_mayThiCongList);
        RecalculateMachineCosts();
    }

    private void DgvVL_CellValueChanged(object sender, DataGridViewCellEventArgs e)
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
        
        // Theo TT37: Hệ số nhiên liệu phụ
        decimal hsXang = 1.02m;
        decimal hsDiezel = 1.03m;
        decimal hsDien = 1.05m;

        _suppressRecalc = true;
        foreach (var m in _mayThiCongList)
        {
            m.NhienLieu = (m.DinhMucXang * gx * hsXang) + 
                          (m.DinhMucDiezel * gdz * hsDiezel) + 
                          (m.DinhMucDien * gdi * hsDien);
        }
        dgvMay.Refresh();
        _suppressRecalc = false;
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
    internal decimal DinhMucXang { get; set; }
    internal decimal DinhMucDiezel { get; set; }
    internal decimal DinhMucDien { get; set; }
}
