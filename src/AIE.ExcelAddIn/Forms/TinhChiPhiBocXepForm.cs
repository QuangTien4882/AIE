using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.ExcelAddIn.Services;

namespace AIE.ExcelAddIn.Forms;

public class TinhChiPhiBocXepForm : Form
{
    private static readonly CultureInfo ViVn = new CultureInfo("vi-VN");

    private readonly string _tenVatLieu;
    private readonly string _donViVatLieu;
    private decimal _donGiaNC;
    private decimal _donGiaMay;
    private readonly string? _maDinhMucCu;
    private readonly int _phamViCu;
    private readonly decimal _dmNCCu;
    private readonly decimal _dmMayCu;

    public decimal KetQuaChiPhiBocXep { get; private set; }
    public DinhMucBocXepItem? SelectedDinhMuc { get; private set; }
    public int SelectedPhamVi { get; private set; } // 0: Cả hai, 1: Lên, 2: Xuống
    public decimal DmNC { get; private set; }
    public decimal DmMay { get; private set; }
    public string? MaMay { get; private set; }

    // UI Controls
    private ComboBox cbDinhMuc;
    private RadioButton rbCaHai;
    private RadioButton rbLen;
    private RadioButton rbXuong;

    private TextBox txtDmNC;
    private TextBox txtDonGiaNC;
    private Label lblThanhTienNC;

    private Label lblTenMay;
    private TextBox txtDmMay;
    private Label lblDvtMay;
    private TextBox txtDonGiaMay;
    private Label lblThanhTienMay;

    private Label lblQuyDoi;
    private Label lblKetQua;

    private bool _suppressRecalc = false;

    public TinhChiPhiBocXepForm(
        string tenVatLieu, 
        string donViVatLieu, 
        decimal donGiaNC, 
        decimal donGiaMay = 0, 
        decimal giaTriHienTai = 0,
        string? maDinhMucCu = null,
        int phamViCu = 0,
        decimal dmNCCu = 0,
        decimal dmMayCu = 0)
    {
        _tenVatLieu = tenVatLieu;
        _donViVatLieu = string.IsNullOrWhiteSpace(donViVatLieu) ? "Đơn vị" : donViVatLieu.Trim();
        _donGiaNC = donGiaNC > 0 ? donGiaNC : 254498m;
        _donGiaMay = donGiaMay > 0 ? donGiaMay : 1858292m;
        _maDinhMucCu = maDinhMucCu;
        _phamViCu = phamViCu;
        _dmNCCu = dmNCCu;
        _dmMayCu = dmMayCu;

        InitializeComponent();
        LoadInitialData();

        if (giaTriHienTai > 0 && KetQuaChiPhiBocXep == 0)
        {
            KetQuaChiPhiBocXep = giaTriHienTai;
        }
    }

    private void InitializeComponent()
    {
        this.Text = $"Tính chi phí bốc xếp - {_tenVatLieu}";
        // Kích thước mặc định rộng rãi 940x650, hỗ trợ thay đổi kích thước linh hoạt theo màn hình
        this.ClientSize = new Size(940, 650);
        this.MinimumSize = new Size(900, 620);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.Sizable; // Cho phép co giãn cửa sổ linh hoạt
        this.MaximizeBox = true;                       // Cho phép phóng to tối đa
        this.MinimizeBox = false;
        this.Font = new Font("Be Vietnam Pro", 9.5f);
        this.BackColor = Color.FromArgb(248, 249, 250);

        // 1. Header Panel (Dock Top)
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 82,
            BackColor = Color.White
        };
        pnlHeader.Paint += (s, e) =>
        {
            e.Graphics.DrawLine(new Pen(Color.FromArgb(222, 226, 230), 1), 0, pnlHeader.Height - 1, pnlHeader.Width, pnlHeader.Height - 1);
        };

        var lblTitle = new Label
        {
            Text = "TÍNH CHI PHÍ BỐC XẾP VẬT LIỆU",
            Font = new Font("Be Vietnam Pro", 12.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215),
            AutoSize = true,
            Location = new Point(25, 14)
        };
        var lblSub = new Label
        {
            Text = $"Vật liệu: {_tenVatLieu}    |    Đơn vị tính: {_donViVatLieu}",
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125),
            AutoSize = true,
            Location = new Point(25, 46)
        };
        pnlHeader.Controls.Add(lblTitle);
        pnlHeader.Controls.Add(lblSub);
        this.Controls.Add(pnlHeader);

        // 2. Bottom Action Panel (Dock Bottom)
        var pnlBottom = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 65,
            BackColor = Color.White
        };
        pnlBottom.Paint += (s, e) =>
        {
            e.Graphics.DrawLine(new Pen(Color.FromArgb(222, 226, 230), 1), 0, 0, pnlBottom.Width, 0);
        };

        var btnApDung = new Button
        {
            Text = "✔  Áp dụng",
            DialogResult = DialogResult.OK,
            Width = 145,
            Height = 40,
            Location = new Point(pnlBottom.Width - 275, 12),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnApDung.FlatAppearance.BorderSize = 0;
        btnApDung.Click += (s, e) =>
        {
            // Lưu lại cấu hình bốc xếp vào kho lưu trữ bền vững
            BocXepStorage.SaveConfig(_tenVatLieu, SelectedDinhMuc?.MaHieu, SelectedPhamVi, DmNC, DmMay, MaMay, KetQuaChiPhiBocXep);
            this.DialogResult = DialogResult.OK;
            this.Close();
        };

        var btnDong = new Button
        {
            Text = "Đóng",
            DialogResult = DialogResult.Cancel,
            Width = 105,
            Height = 40,
            Location = new Point(pnlBottom.Width - 120, 12),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Be Vietnam Pro", 9.5f),
            Cursor = Cursors.Hand
        };
        btnDong.FlatAppearance.BorderColor = Color.FromArgb(206, 212, 218);

        pnlBottom.Controls.Add(btnApDung);
        pnlBottom.Controls.Add(btnDong);
        this.Controls.Add(pnlBottom);

        // 3. Body Content Panel (Dock Fill với AutoScroll)
        var pnlContent = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(25, 14, 25, 14)
        };
        this.Controls.Add(pnlContent);
        pnlContent.BringToFront();

        int y = 14;

        // Group 1: Vị trí chọn định mức áp dụng (Tự động co giãn Anchor Left | Right)
        var gbDinhMuc = new GroupBox
        {
            Text = "Định mức bốc xếp áp dụng (theo Chương XII - Thông tư số 38/2026/TT-BXD)",
            Location = new Point(25, y),
            Size = new Size(875, 72),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        cbDinhMuc = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Location = new Point(15, 27),
            Size = new Size(845, 28),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Regular),
            DropDownWidth = 900
        };
        foreach (var item in DinhMucBocXepDatabase.DanhSach)
        {
            cbDinhMuc.Items.Add(item);
        }
        cbDinhMuc.SelectedIndexChanged += (s, e) => OnDinhMucChanged();
        gbDinhMuc.Controls.Add(cbDinhMuc);
        pnlContent.Controls.Add(gbDinhMuc);
        y += 82;

        // Group 2: Phạm vi bốc xếp (Anchor Left | Right)
        var gbPhamVi = new GroupBox
        {
            Text = "Phạm vi công tác bốc xếp",
            Location = new Point(25, y),
            Size = new Size(875, 65),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Be Vietnam Pro", 9.5f),
            ForeColor = Color.FromArgb(33, 37, 41)
        };
        rbCaHai = new RadioButton { Text = "Cả bốc lên và bốc xuống", Location = new Point(25, 26), Size = new Size(240, 26), Checked = true };
        rbLen = new RadioButton { Text = "Chỉ bốc lên", Location = new Point(310, 26), Size = new Size(180, 26) };
        rbXuong = new RadioButton { Text = "Chỉ bốc xuống", Location = new Point(530, 26), Size = new Size(180, 26) };

        rbCaHai.CheckedChanged += (s, e) => { if (rbCaHai.Checked) UpdateDinhMucValues(); };
        rbLen.CheckedChanged += (s, e) => { if (rbLen.Checked) UpdateDinhMucValues(); };
        rbXuong.CheckedChanged += (s, e) => { if (rbXuong.Checked) UpdateDinhMucValues(); };

        gbPhamVi.Controls.Add(rbCaHai);
        gbPhamVi.Controls.Add(rbLen);
        gbPhamVi.Controls.Add(rbXuong);
        pnlContent.Controls.Add(gbPhamVi);
        y += 75;

        // Group 3: Chi tiết hao phí và đơn giá (Anchor Left | Right, Cột Thành tiền rộng 190px đảm bảo 1 dòng)
        var gbChiTiet = new GroupBox
        {
            Text = "Chi tiết định mức và chiết tính đơn giá",
            Location = new Point(25, y),
            Size = new Size(875, 165),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Be Vietnam Pro", 9.5f),
            ForeColor = Color.FromArgb(33, 37, 41)
        };

        var tblGrid = new TableLayoutPanel
        {
            Location = new Point(15, 25),
            Size = new Size(845, 128),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            ColumnCount = 5,
            RowCount = 3
        };
        // Cột 0: Tự động co giãn theo chiều rộng cửa sổ; Các cột số liệu có độ rộng tuyệt đối chuẩn mực
        tblGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // Tên thành phần: co giãn linh hoạt
        tblGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f)); // Định mức: 100px
        tblGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));  // ĐVT: 70px
        tblGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f)); // Đơn giá: 150px
        tblGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190f)); // Thành tiền: 190px (rộng rãi, 100% trọn vẹn trên 1 dòng)

        tblGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 36f)); // Header row
        tblGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f)); // Row 1 (NC)
        tblGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f)); // Row 2 (Máy)

        // Tiêu đề cột
        var h1 = new Label { Text = "Thành phần hao phí", Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(70, 80, 95), AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Bottom, Margin = new Padding(5, 0, 0, 8) };
        var h2 = new Label { Text = "Định mức", Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(70, 80, 95), AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Bottom, Margin = new Padding(5, 0, 0, 8) };
        var h3 = new Label { Text = "ĐVT", Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(70, 80, 95), TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8) };
        var h4 = new Label { Text = "Đơn giá (đồng)", Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(70, 80, 95), AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Bottom, Margin = new Padding(5, 0, 0, 8) };
        var h5 = new Label { Text = "Thành tiền (đồng)", Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(70, 80, 95), AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Bottom, Margin = new Padding(5, 0, 0, 8) };
        tblGrid.Controls.Add(h1, 0, 0);
        tblGrid.Controls.Add(h2, 1, 0);
        tblGrid.Controls.Add(h3, 2, 0);
        tblGrid.Controls.Add(h4, 3, 0);
        tblGrid.Controls.Add(h5, 4, 0);

        // Row 1: Nhân công (Căn chỉnh baseline ngang và dọc chuẩn xác)
        var lblNCName = new Label { Text = "1. Nhân công nhóm I", Font = new Font("Be Vietnam Pro", 9.5f), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDmNC = new TextBox { Width = 85, Height = 26, TextAlign = HorizontalAlignment.Right, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDmNC.TextChanged += (s, e) => TinhToan();
        var lblDvtNC = new Label { Text = "công", Font = new Font("Be Vietnam Pro", 9.5f), TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(0) };
        txtDonGiaNC = new TextBox { Width = 130, Height = 26, TextAlign = HorizontalAlignment.Right, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDonGiaNC.TextChanged += (s, e) => TinhToan();
        lblThanhTienNC = new Label { Text = "0 đ", Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold), ForeColor = Color.FromArgb(40, 167, 69), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };

        tblGrid.Controls.Add(lblNCName, 0, 1);
        tblGrid.Controls.Add(txtDmNC, 1, 1);
        tblGrid.Controls.Add(lblDvtNC, 2, 1);
        tblGrid.Controls.Add(txtDonGiaNC, 3, 1);
        tblGrid.Controls.Add(lblThanhTienNC, 4, 1);

        // Row 2: Máy thi công
        lblTenMay = new Label { Text = "2. Cần cẩu bánh hơi - sức nâng: 6 T", Font = new Font("Be Vietnam Pro", 9.5f), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDmMay = new TextBox { Width = 85, Height = 26, TextAlign = HorizontalAlignment.Right, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDmMay.TextChanged += (s, e) => TinhToan();
        lblDvtMay = new Label { Text = "ca", Font = new Font("Be Vietnam Pro", 9.5f), TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, Margin = new Padding(0) };
        txtDonGiaMay = new TextBox { Width = 130, Height = 26, TextAlign = HorizontalAlignment.Right, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };
        txtDonGiaMay.TextChanged += (s, e) => TinhToan();
        lblThanhTienMay = new Label { Text = "0 đ", Font = new Font("Be Vietnam Pro", 10f, FontStyle.Bold), ForeColor = Color.FromArgb(40, 167, 69), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(5, 0, 0, 0) };

        tblGrid.Controls.Add(lblTenMay, 0, 2);
        tblGrid.Controls.Add(txtDmMay, 1, 2);
        tblGrid.Controls.Add(lblDvtMay, 2, 2);
        tblGrid.Controls.Add(txtDonGiaMay, 3, 2);
        tblGrid.Controls.Add(lblThanhTienMay, 4, 2);

        gbChiTiet.Controls.Add(tblGrid);
        pnlContent.Controls.Add(gbChiTiet);
        y += 175;

        // Group 4: Quy đổi & Kết quả (Anchor Left | Right)
        var pnlKetQua = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(875, 80),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.FromArgb(235, 245, 255),
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(15, 8, 15, 8)
        };

        lblQuyDoi = new Label
        {
            Text = "Quy đổi đơn vị tính: ...",
            Location = new Point(15, 8),
            AutoSize = true,
            Font = new Font("Be Vietnam Pro", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(70, 80, 95)
        };

        lblKetQua = new Label
        {
            Text = "Chi phí bốc xếp: 0 đ / đơn vị",
            Location = new Point(15, 36),
            AutoSize = true,
            Font = new Font("Be Vietnam Pro", 13.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };

        pnlKetQua.Controls.Add(lblQuyDoi);
        pnlKetQua.Controls.Add(lblKetQua);
        pnlContent.Controls.Add(pnlKetQua);
    }

    private void LoadInitialData()
    {
        // 1. Chọn định mức: ưu tiên mã định mức cũ đã cấu hình trước đó, hoặc tìm trong kho lưu trữ bền vững
        string? targetMa = _maDinhMucCu;
        if (string.IsNullOrEmpty(targetMa))
        {
            var savedCfg = BocXepStorage.GetConfig(_tenVatLieu);
            targetMa = savedCfg?.MaDinhMuc;
        }

        if (string.IsNullOrEmpty(targetMa))
        {
            var matched = DinhMucBocXepDatabase.NhanDienDinhMuc(_tenVatLieu, _donViVatLieu);
            targetMa = matched?.MaHieu;
        }

        if (!string.IsNullOrEmpty(targetMa))
        {
            var idx = cbDinhMuc.Items.Cast<DinhMucBocXepItem>().ToList().FindIndex(x => x.MaHieu == targetMa);
            cbDinhMuc.SelectedIndex = idx >= 0 ? idx : 0;
        }
        else
        {
            cbDinhMuc.SelectedIndex = 0;
        }

        // 2. Phạm vi bốc xếp cũ
        int phamVi = _phamViCu;
        if (phamVi == 0)
        {
            var savedCfg = BocXepStorage.GetConfig(_tenVatLieu);
            if (savedCfg != null) phamVi = savedCfg.PhamVi;
        }

        if (phamVi == 1) rbLen.Checked = true;
        else if (phamVi == 2) rbXuong.Checked = true;
        else rbCaHai.Checked = true;

        txtDonGiaNC.Text = _donGiaNC.ToString("N0", ViVn);
        txtDonGiaMay.Text = _donGiaMay.ToString("N0", ViVn);

        UpdateDinhMucValues();

        // 3. Nếu trước đó có định mức tùy chỉnh tay:
        decimal dmNC = _dmNCCu;
        decimal dmMay = _dmMayCu;
        if (dmNC <= 0 && dmMay <= 0)
        {
            var savedCfg = BocXepStorage.GetConfig(_tenVatLieu);
            if (savedCfg != null)
            {
                dmNC = savedCfg.DmNC;
                dmMay = savedCfg.DmMay;
            }
        }

        if (dmNC > 0) txtDmNC.Text = dmNC.ToString("0.###", ViVn);
        if (dmMay > 0) txtDmMay.Text = dmMay.ToString("0.###", ViVn);
    }

    private void OnDinhMucChanged()
    {
        if (cbDinhMuc.SelectedItem is not DinhMucBocXepItem dm) return;
        SelectedDinhMuc = dm;

        bool coTachLenXuong = dm.DmNCLen > 0 && dm.DmNCXuong > 0;
        rbLen.Enabled = coTachLenXuong || dm.DmNCLen > 0;
        rbXuong.Enabled = coTachLenXuong || dm.DmNCXuong > 0;

        lblTenMay.Visible = dm.CoMay;
        txtDmMay.Visible = dm.CoMay;
        lblDvtMay.Visible = dm.CoMay;
        txtDonGiaMay.Visible = dm.CoMay;
        lblThanhTienMay.Visible = dm.CoMay;

        if (dm.CoMay)
        {
            lblTenMay.Text = $"2. {dm.TenMay}";
        }

        UpdateDinhMucValues();
    }

    private void UpdateDinhMucValues()
    {
        if (cbDinhMuc.SelectedItem is not DinhMucBocXepItem dm) return;

        _suppressRecalc = true;

        decimal dmNC = 0;
        decimal dmMay = 0;

        if (rbLen.Checked)
        {
            SelectedPhamVi = 1;
            dmNC = dm.DmNCLen;
            dmMay = dm.DmMayLen;
        }
        else if (rbXuong.Checked)
        {
            SelectedPhamVi = 2;
            dmNC = dm.DmNCXuong;
            dmMay = dm.DmMayXuong;
        }
        else // Cả hai
        {
            SelectedPhamVi = 0;
            dmNC = dm.DmNCCaHai > 0 ? dm.DmNCCaHai : (dm.DmNCLen + dm.DmNCXuong);
            dmMay = dm.DmMayCaHai > 0 ? dm.DmMayCaHai : (dm.DmMayLen + dm.DmMayXuong);
        }

        txtDmNC.Text = dmNC.ToString("0.###", ViVn);
        txtDmMay.Text = dmMay.ToString("0.###", ViVn);

        _suppressRecalc = false;
        TinhToan();
    }

    private void TinhToan()
    {
        if (_suppressRecalc) return;
        if (cbDinhMuc.SelectedItem is not DinhMucBocXepItem dm) return;

        decimal dmNC = ParseVn(txtDmNC.Text);
        decimal dgNC = ParseVn(txtDonGiaNC.Text);
        decimal ttNC = dmNC * dgNC;
        lblThanhTienNC.Text = $"{ttNC.ToString("N0", ViVn)} đ";

        decimal ttMay = 0;
        decimal dmMay = 0;
        if (dm.CoMay)
        {
            dmMay = ParseVn(txtDmMay.Text);
            decimal dgMay = ParseVn(txtDonGiaMay.Text);
            ttMay = dmMay * dgMay;
            lblThanhTienMay.Text = $"{ttMay.ToString("N0", ViVn)} đ";
        }

        decimal tongChiPhi1Dm = ttNC + ttMay;

        decimal heSoQuyDoi = DinhMucBocXepDatabase.TinhHeSoQuyDoi(_donViVatLieu, dm.DonViDinhMuc);
        KetQuaChiPhiBocXep = Math.Round(tongChiPhi1Dm * heSoQuyDoi, 0, MidpointRounding.AwayFromZero);

        // Lưu lại các thông số cấu hình bốc xếp
        DmNC = dmNC;
        DmMay = dmMay;
        MaMay = dm.CoMay ? dm.MaMay : null;

        if (heSoQuyDoi != 1.0m)
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dm.DonViDinhMuc} = {tongChiPhi1Dm.ToString("N0", ViVn)} đ   (Quy đổi về 1 {_donViVatLieu} = {heSoQuyDoi:0.####} {dm.DonViDinhMuc})";
        }
        else
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dm.DonViDinhMuc}   (Cùng đơn vị tính với vật tư)";
        }

        lblKetQua.Text = $"Chi phí bốc xếp: {KetQuaChiPhiBocXep.ToString("N0", ViVn)} đ / {_donViVatLieu}";
    }

    private static decimal ParseVn(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        var clean = text.Trim().Replace(".", "").Replace(",", ".");
        return decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0;
    }
}
