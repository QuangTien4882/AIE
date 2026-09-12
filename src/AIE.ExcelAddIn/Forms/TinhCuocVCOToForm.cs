using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.ExcelAddIn.Services;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms;

public class TinhCuocVCOToForm : Form
{
    private static readonly CultureInfo ViVn = new CultureInfo("vi-VN")
    {
        NumberFormat =
        {
            NumberDecimalSeparator = ",",
            NumberGroupSeparator = "."
        }
    };

    private readonly string _tenVatLieu;
    private readonly string _donViVatLieu;
    private decimal _donGiaCaMay;
    private readonly string? _maMayCu;
    private readonly string? _maDinhMucCu;
    private readonly Func<string, string, decimal>? _layGiaMayFunc;

    public decimal KetQuaCuocOTo { get; private set; }
    public DinhMucVCOToItem? SelectedDinhMuc { get; private set; }
    public string MaMay { get; private set; } = string.Empty;
    public decimal DonGiaCaMay { get; private set; }
    public decimal TongCuLyKm { get; private set; } = 0;
    public List<CungDuongVanChuyen> CungDuongs { get; private set; } = new();

    // UI Controls
    private Label lblCuLy;
    private TextBox txtTongCuLy;
    private Label lblLoaiVL;
    private ComboBox cbVatLieu;
    private Label lblGiaMay;
    private TextBox txtDonGiaCaMay;
    private Label lblXe;
    private ComboBox cbLoaiXe;
    private Label lblDmInfo;

    private DataGridView dgvCungDuong;
    private ComboBox cbTemplates;
    private Button btnLuuTemplate;
    private Button btnXoaTemplate;
    private Button btnThemDoan;
    private Button btnChenDoan;
    private Button btnXoaDoan;
    private Button btnTraCuuDuong;

    private Panel pnlChietTinh;
    private Label lblHuongDan;
    private Label lblNac1;
    private Label lblNac2;
    private Label lblNac3;
    private Label lblNac4;
    private Label lblTongCaXe;

    private Label lblQuyDoi;
    private Label lblKetQua;

    private bool _suppressRecalc = false;
    private bool _isUpdatingGridFromCode = false;

    public TinhCuocVCOToForm(
        string tenVatLieu, 
        string donViVatLieu, 
        decimal donGiaCaMay = 0, 
        decimal giaTriHienTai = 0,
        string? maDinhMucCu = null,
        string? maMayCu = null,
        Func<string, string, decimal>? layGiaMayFunc = null)
    {
        _tenVatLieu = tenVatLieu;
        _donViVatLieu = string.IsNullOrWhiteSpace(donViVatLieu) ? "ĐVT" : donViVatLieu.Trim();
        _donGiaCaMay = donGiaCaMay > 0 ? donGiaCaMay : 1040324m;
        _maDinhMucCu = maDinhMucCu;
        _maMayCu = maMayCu;
        _layGiaMayFunc = layGiaMayFunc;

        InitializeComponent();
        FormStateHelper.Attach(this);
        LoadInitialData();

        if (giaTriHienTai > 0 && KetQuaCuocOTo == 0)
        {
            KetQuaCuocOTo = giaTriHienTai;
        }
    }

    private void InitializeComponent()
    {
        this.Text = $"Tính cước vận chuyển bằng ô tô - {_tenVatLieu}";
        
        var area = Screen.PrimaryScreen.WorkingArea;
        int targetW = Math.Min(1200, area.Width - 40);
        int targetH = Math.Min(880, area.Height - 40);
        this.ClientSize = new Size(targetW, targetH);
        int minW = Math.Min(980, area.Width - 40);
        int minH = Math.Min(580, area.Height - 40);
        this.MinimumSize = new Size(minW, minH);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.Sizable;
        this.MaximizeBox = true;
        this.MinimizeBox = false;
        this.AutoValidate = AutoValidate.Disable;
        this.Font = new Font("Segoe UI", 9.5f);
        this.BackColor = Color.FromArgb(248, 249, 250);

        // 1. Header Panel
        var pnlHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 76,
            BackColor = Color.White
        };
        pnlHeader.Paint += (s, e) =>
        {
            e.Graphics.DrawLine(new Pen(Color.FromArgb(222, 226, 230), 1), 0, pnlHeader.Height - 1, pnlHeader.Width, pnlHeader.Height - 1);
        };

        var lblTitle = new Label
        {
            Text = "TÍNH CƯỚC VẬN CHUYỂN VẬT LIỆU BẰNG Ô TÔ",
            Font = new Font("Segoe UI", 13.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 120, 215),
            AutoSize = true,
            Location = new Point(25, 12)
        };
        var lblSub = new Label
        {
            Text = $"Vật liệu: {_tenVatLieu}    |    Đơn vị tính: {_donViVatLieu}    |    Áp dụng theo Chương XII - Thông tư số 38/2026/TT-BXD",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125),
            AutoSize = true,
            Location = new Point(25, 44)
        };
        pnlHeader.Controls.Add(lblTitle);
        pnlHeader.Controls.Add(lblSub);
        this.Controls.Add(pnlHeader);

        // 2. Bottom Action Panel
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
            Width = 145,
            Height = 42,
            Location = new Point(pnlBottom.Width - 275, 11),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnApDung.FlatAppearance.BorderSize = 0;
        btnApDung.Click += (s, e) =>
        {
            ReadCungDuongsFromGrid();
            VanChuyenStorage.SaveConfigOTo(
                _tenVatLieu, 
                SelectedDinhMuc?.MaHieu ?? "", 
                MaMay, 
                DonGiaCaMay, 
                CungDuongs, 
                KetQuaCuocOTo);
            this.DialogResult = DialogResult.OK;
            this.Close();
        };

        var btnDong = new Button
        {
            Text = "Đóng",
            Width = 110,
            Height = 42,
            Location = new Point(pnlBottom.Width - 120, 11),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f),
            Cursor = Cursors.Hand
        };
        btnDong.FlatAppearance.BorderColor = Color.FromArgb(206, 212, 218);
        btnDong.Click += (s, e) =>
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        };
        this.CancelButton = btnDong;

        pnlBottom.Controls.Add(btnApDung);
        pnlBottom.Controls.Add(btnDong);
        this.Controls.Add(pnlBottom);

        // 3. Body Content Panel
        var pnlContent = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = Color.FromArgb(248, 249, 250),
            Padding = new Padding(25, 15, 25, 15)
        };
        this.Controls.Add(pnlContent);
        pnlContent.BringToFront();

        int y = 10;

        // =========================================================================
        // CARD 1: THÔNG TIN CHUNG & PHƯƠNG TIỆN (BỐ CỤC THÔNG MINH, KHÔNG BAO GIỜ TRÀN)
        // =========================================================================
        var cardDinhMuc = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 160),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.White
        };
        cardDinhMuc.Paint += (s, e) =>
        {
            var g = e.Graphics;
            var r = cardDinhMuc.ClientRectangle;
            g.DrawRectangle(new Pen(Color.FromArgb(222, 226, 230), 1), 0, 0, r.Width - 1, r.Height - 1);
            g.FillRectangle(new SolidBrush(Color.FromArgb(243, 248, 255)), 1, 1, r.Width - 2, 34);
            g.DrawLine(new Pen(Color.FromArgb(215, 230, 250), 1), 1, 35, r.Width - 2, 35);
        };

        var lblCard1Title = new Label
        {
            Text = "📋  THÔNG TIN CHUNG & PHƯƠNG TIỆN VẬN CHUYỂN (Chương XII - Thông tư số 38/2026/TT-BXD)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 9),
            BackColor = Color.Transparent
        };
        cardDinhMuc.Controls.Add(lblCard1Title);

        // HÀNG 1: TỔNG CỰ LY | NHÓM VẬT LIỆU | GIÁ CA XE
        lblCuLy = new Label { Text = "Tổng cự ly vận chuyển (km):", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
        txtTongCuLy = new TextBox
        {
            Size = new Size(115, 26),
            TextAlign = HorizontalAlignment.Right,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };
        txtTongCuLy.KeyDown += (s, e) => 
        { 
            if (e.KeyCode == Keys.Enter) 
            { 
                OnTongCuLyUserSubmitted(); 
                e.SuppressKeyPress = true; 
            } 
        };
        txtTongCuLy.Leave += (s, e) => OnTongCuLyUserSubmitted();

        lblLoaiVL = new Label { Text = "Nhóm vật liệu:", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular) };
        cbVatLieu = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Size = new Size(200, 26),
            DropDownWidth = 280,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };
        var nhomVLs = DinhMucVanChuyenDatabase.DanhSachOTo.Select(x => x.LoaiVatLieu).Distinct().ToList();
        foreach (var n in nhomVLs) cbVatLieu.Items.Add(n);
        cbVatLieu.SelectedIndexChanged += (s, e) => OnNhomVatLieuChanged();

        lblGiaMay = new Label { Text = "Giá ca xe (đ):", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular) };
        txtDonGiaCaMay = new TextBox
        {
            Size = new Size(140, 26),
            TextAlign = HorizontalAlignment.Right,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold)
        };
        txtDonGiaCaMay.TextChanged += (s, e) => TinhToan();

        // HÀNG 2: PHƯƠNG TIỆN VẬN CHUYỂN
        lblXe = new Label { Text = "Phương tiện vận chuyển:", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular) };
        cbLoaiXe = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Height = 26,
            DropDownWidth = 660,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };
        cbLoaiXe.SelectedIndexChanged += (s, e) => OnLoaiXeChanged();

        // HÀNG 3: THÔNG TIN ĐỊNH MỨC TÓM TẮT
        lblDmInfo = new Label
        {
            Text = "Nhập các đoạn theo hành trình thực tế từ nguồn cung đến công trình để phần mềm tự động phân bổ",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Italic),
            ForeColor = Color.FromArgb(0, 102, 204)
        };

        cardDinhMuc.Controls.AddRange(new Control[] { lblCuLy, txtTongCuLy, lblLoaiVL, cbVatLieu, lblGiaMay, txtDonGiaCaMay, lblXe, cbLoaiXe, lblDmInfo });

        Action layoutCard1 = () =>
        {
            // Hàng 1
            lblCuLy.Location = new Point(16, 49);
            txtTongCuLy.Location = new Point(lblCuLy.Right + 10, 45);

            lblLoaiVL.Location = new Point(txtTongCuLy.Right + 25, 49);
            cbVatLieu.Location = new Point(lblLoaiVL.Right + 10, 45);

            lblGiaMay.Location = new Point(cbVatLieu.Right + 25, 49);
            txtDonGiaCaMay.Location = new Point(lblGiaMay.Right + 10, 45);

            // Hàng 2: cbLoaiXe tự co giãn theo chiều rộng của Card
            lblXe.Location = new Point(16, 84);
            int cbLeft = lblXe.Right + 10;
            int cbWidth = Math.Max(300, cardDinhMuc.ClientSize.Width - cbLeft - 25);
            cbLoaiXe.Location = new Point(cbLeft, 80);
            cbLoaiXe.Size = new Size(cbWidth, 26);

            // Hàng 3
            lblDmInfo.Location = new Point(16, 120);
        };
        cardDinhMuc.Layout += (s, e) => layoutCard1();
        cardDinhMuc.Resize += (s, e) => layoutCard1();
        layoutCard1();

        pnlContent.Controls.Add(cardDinhMuc);
        y += 175;

        // =========================================================================
        // CARD 2: CHI TIẾT CÁC ĐOẠN ĐƯỜNG THEO HÀNH TRÌNH THỰC TẾ (THÔNG TƯ 38)
        // =========================================================================
        var cardCungDuong = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 325),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.White
        };
        cardCungDuong.Paint += (s, e) =>
        {
            var g = e.Graphics;
            var r = cardCungDuong.ClientRectangle;
            g.DrawRectangle(new Pen(Color.FromArgb(222, 226, 230), 1), 0, 0, r.Width - 1, r.Height - 1);
            g.FillRectangle(new SolidBrush(Color.FromArgb(243, 248, 255)), 1, 1, r.Width - 2, 34);
            g.DrawLine(new Pen(Color.FromArgb(215, 230, 250), 1), 1, 35, r.Width - 2, 35);
        };

        var lblCard2Title = new Label
        {
            Text = "🛣️  CHI TIẾT CÁC ĐOẠN ĐƯỜNG VẬN CHUYỂN (Nhập các đoạn theo hành trình thực tế từ nơi cung ứng đến công trình)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 9),
            BackColor = Color.Transparent
        };
        cardCungDuong.Controls.Add(lblCard2Title);

        // Thanh công cụ mẫu tuyến đường (Dùng FlowLayoutPanel để các nút không bao giờ bị đè lấn nhau)
        var pnlTemplateBar = new FlowLayoutPanel
        {
            Location = new Point(16, 42),
            Size = new Size(cardCungDuong.ClientSize.Width - 32, 34),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.Transparent,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoScroll = false
        };

        var lblTuyenMau = new Label
        {
            Text = "Mẫu tuyến đường:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Margin = new Padding(0, 5, 8, 0)
        };

        cbTemplates = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Width = 280,
            DropDownWidth = 450,
            Margin = new Padding(0, 1, 10, 0)
        };
        cbTemplates.SelectedIndexChanged += (s, e) =>
        {
            if (cbTemplates.SelectedItem is TuyenDuongTemplate t)
            {
                if (t.CungDuongs != null && t.CungDuongs.Count > 0)
                {
                    bool needConfirm = dgvCungDuong.Rows.Count > 0;
                    if (needConfirm)
                    {
                        var confirm = MessageBox.Show(
                            $"Áp dụng mẫu tuyến \"{t.TenTemplate}\" ({t.CungDuongs.Count} đoạn, cự ly {t.TongCuLyKm:0.###} km)?\nCác đoạn đường hiện tại trong bảng sẽ được thay thế.",
                            "Áp dụng mẫu tuyến đường",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
                        if (confirm != DialogResult.Yes) return;
                    }
                    ApDungTemplate(t);
                }
            }
        };

        btnLuuTemplate = new Button
        {
            Text = "💾 Lưu mẫu tuyến...",
            Size = new Size(145, 28),
            Margin = new Padding(0, 0, 8, 0),
            BackColor = Color.FromArgb(23, 162, 184),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnLuuTemplate.FlatAppearance.BorderSize = 0;
        btnLuuTemplate.Click += (s, e) => OnLuuTemplateClicked();

        btnXoaTemplate = new Button
        {
            Text = "🗑 Xóa mẫu",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 0, 0),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f),
            ForeColor = Color.FromArgb(220, 53, 69),
            Cursor = Cursors.Hand
        };
        btnXoaTemplate.FlatAppearance.BorderColor = Color.FromArgb(220, 53, 69);
        btnXoaTemplate.Click += (s, e) => OnXoaTemplateClicked();

        pnlTemplateBar.Controls.AddRange(new Control[] { lblTuyenMau, cbTemplates, btnLuuTemplate, btnXoaTemplate });
        cardCungDuong.Controls.Add(pnlTemplateBar);

        dgvCungDuong = new DataGridView
        {
            Location = new Point(16, 78),
            Size = new Size(cardCungDuong.ClientSize.Width - 32, 182),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.Fixed3D,
            RowHeadersVisible = false,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            MultiSelect = false,
            EditMode = DataGridViewEditMode.EditOnEnter
        };
        dgvCungDuong.RowTemplate.Height = 29;

        // Click 1 lần là edit được ngay, nếu là dropdown thì bung danh sách ngay lập tức
        dgvCungDuong.CellClick += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgvCungDuong.BeginEdit(true);
                if (dgvCungDuong.EditingControl is ComboBox cb)
                {
                    cb.DroppedDown = true;
                }
            }
        };

        // Bắt phím Enter trong DataGridView để không nhảy ra sheet Excel
        dgvCungDuong.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvCungDuong.EndEdit();
                e.Handled = true;
            }
        };

        var colStt = new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 45, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
        var colDiemDau = new DataGridViewTextBoxColumn { Name = "DiemDau", HeaderText = "Điểm đầu", Width = 150 };
        var colDiemCuoi = new DataGridViewTextBoxColumn { Name = "DiemCuoi", HeaderText = "Điểm cuối", Width = 150 };
        var colTen = new DataGridViewTextBoxColumn { Name = "TenDoan", HeaderText = "Tên tuyến đường", Width = 180 };
        var colCuLy = new DataGridViewTextBoxColumn { Name = "CuLy", HeaderText = "Cự ly (km)", Width = 110, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
        var colLoai = new DataGridViewComboBoxColumn { Name = "LoaiDuong", HeaderText = "Loại đường", Width = 160 };
        colLoai.Items.AddRange("Loại 1 (k=0,57)", "Loại 2 (k=0,68)", "Loại 3 (k=1,00)", "Loại 4 (k=1,35)", "Loại 5 (k=1,50)", "Loại 6 (k=1,80)");
        var colHeSo = new DataGridViewTextBoxColumn { Name = "HeSoK", HeaderText = "Hệ số kđ", Width = 85, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) } };
        var colGhiChu = new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill };

        dgvCungDuong.Columns.AddRange(colStt, colDiemDau, colDiemCuoi, colTen, colCuLy, colLoai, colHeSo, colGhiChu);
        dgvCungDuong.CellValueChanged += DgvCungDuong_CellValueChanged;
        dgvCungDuong.CellEndEdit += (s, e) => 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var colName = dgvCungDuong.Columns[e.ColumnIndex].Name;
                if (colName == "CuLy")
                {
                    decimal val = ParseVnFlex(dgvCungDuong["CuLy", e.RowIndex].Value?.ToString());
                    dgvCungDuong["CuLy", e.RowIndex].Value = val > 0 ? val.ToString("0.###", ViVn) : "0";
                }
            }
            if (!_isUpdatingGridFromCode)
            {
                CapNhatTongCuLyTuBang();
            }
            TinhToan();
        };

        dgvCungDuong.CellFormatting += (s, e) =>
        {
            if (e.RowIndex >= 0 && dgvCungDuong.Columns[e.ColumnIndex].Name == "CuLy" && e.Value != null)
            {
                decimal v = ParseVnFlex(e.Value.ToString());
                e.Value = v > 0 ? v.ToString("0.###", ViVn) : "0";
                e.FormattingApplied = true;
            }
        };

        // Thanh nút bấm tinh gọn, đúng chức năng
        var pnlButtons = new FlowLayoutPanel
        {
            Location = new Point(16, 268),
            Size = new Size(cardCungDuong.ClientSize.Width - 32, 42),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false,
            AutoScroll = false
        };

        btnThemDoan = new Button
        {
            Text = "+ Thêm đoạn đường tiếp theo",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 0, 14, 0),
            Margin = new Padding(0, 0, 10, 0),
            BackColor = Color.FromArgb(0, 123, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnThemDoan.FlatAppearance.BorderSize = 0;
        btnThemDoan.Click += (s, e) => 
        {
            string diemDauMacDinh = "";
            if (dgvCungDuong.Rows.Count > 0)
            {
                var lastRow = dgvCungDuong.Rows[dgvCungDuong.Rows.Count - 1];
                diemDauMacDinh = lastRow.Cells["DiemCuoi"].Value?.ToString() ?? "";
            }
            ThemCungDuongMoi(diemDauMacDinh, "", $"Đoạn {dgvCungDuong.Rows.Count + 1}", 0, 3, insertIndex: -1);
        };

        btnChenDoan = new Button
        {
            Text = "➕ Chèn đoạn vào vị trí đang chọn",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 0, 14, 0),
            Margin = new Padding(0, 0, 10, 0),
            BackColor = Color.FromArgb(40, 167, 69),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnChenDoan.FlatAppearance.BorderSize = 0;
        btnChenDoan.Click += (s, e) => ChenDoanVaoViTriDangChon();

        btnXoaDoan = new Button
        {
            Text = "- Xóa đoạn chọn",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 0, 14, 0),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnXoaDoan.FlatAppearance.BorderColor = Color.FromArgb(206, 212, 218);
        btnXoaDoan.Click += (s, e) => XoaCungDuongChon();

        btnTraCuuDuong = new Button
        {
            Text = "🔍 Chọn từ danh mục đường ĐN (QĐ 2035)",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 0, 14, 0),
            Margin = new Padding(0, 0, 0, 0),
            BackColor = Color.FromArgb(240, 244, 248),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnTraCuuDuong.FlatAppearance.BorderColor = Color.FromArgb(180, 205, 235);
        btnTraCuuDuong.Click += (s, e) => MoPopupTraCuuDuong();

        pnlButtons.Controls.Add(btnThemDoan);
        pnlButtons.Controls.Add(btnChenDoan);
        pnlButtons.Controls.Add(btnXoaDoan);
        pnlButtons.Controls.Add(btnTraCuuDuong);

        cardCungDuong.Controls.Add(dgvCungDuong);
        cardCungDuong.Controls.Add(pnlButtons);
        pnlContent.Controls.Add(cardCungDuong);
        y += 340;

        // =========================================================================
        // CARD 3: TỔNG HỢP CHIẾT TÍNH THEO THÔNG TƯ SỐ 38/2026/TT-BXD
        // =========================================================================
        pnlChietTinh = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 160),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.White
        };
        pnlChietTinh.Paint += (s, e) =>
        {
            var g = e.Graphics;
            var r = pnlChietTinh.ClientRectangle;
            g.DrawRectangle(new Pen(Color.FromArgb(222, 226, 230), 1), 0, 0, r.Width - 1, r.Height - 1);
            g.FillRectangle(new SolidBrush(Color.FromArgb(243, 248, 255)), 1, 1, r.Width - 2, 32);
            g.DrawLine(new Pen(Color.FromArgb(215, 230, 250), 1), 1, 33, r.Width - 2, 33);
        };

        var lblCard3Title = new Label
        {
            Text = "📊  TỔNG HỢP CHIẾT TÍNH THEO CÁC NẤC ĐỊNH MỨC CỦA THÔNG TƯ SỐ 38/2026/TT-BXD",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 8),
            BackColor = Color.Transparent
        };
        pnlChietTinh.Controls.Add(lblCard3Title);

        lblHuongDan = new Label
        {
            Text = "Vui lòng nhập các đoạn đường ở trên để tính toán chiết tính theo Thông tư số 38/2026/TT-BXD.",
            Location = new Point(20, 45),
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        lblNac1 = new Label { Text = "• Nấc 1 (Phạm vi <= 1km): ...", Location = new Point(20, 42), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };
        lblNac2 = new Label { Text = "• Nấc 2 (Cự ly 1-10km): ...", Location = new Point(20, 68), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };
        lblNac3 = new Label { Text = "• Nấc 3 (Cự ly 10-60km): ...", Location = new Point(20, 94), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };
        lblNac4 = new Label { Text = "• Nấc 4 (Cự ly > 60km): ...", Location = new Point(20, 120), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };

        lblTongCaXe = new Label
        {
            Text = "Tổng hao phí ca máy = 0 ca",
            Location = new Point(20, 134),
            AutoSize = true,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 167, 69)
        };

        pnlChietTinh.Controls.Add(lblHuongDan);
        pnlChietTinh.Controls.Add(lblNac1);
        pnlChietTinh.Controls.Add(lblNac2);
        pnlChietTinh.Controls.Add(lblNac3);
        pnlChietTinh.Controls.Add(lblNac4);
        pnlChietTinh.Controls.Add(lblTongCaXe);
        pnlContent.Controls.Add(pnlChietTinh);
        y += 175;

        // =========================================================================
        // CARD 4: QUY ĐỔI & KẾT QUẢ CƯỚC VẬN CHUYỂN
        // =========================================================================
        var pnlKetQua = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 85),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.FromArgb(235, 245, 255),
            BorderStyle = BorderStyle.FixedSingle,
            Padding = new Padding(20, 12, 20, 12)
        };

        lblQuyDoi = new Label
        {
            Text = "Quy đổi đơn vị tính: ...",
            Location = new Point(18, 12),
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(70, 80, 95)
        };

        lblKetQua = new Label
        {
            Text = "Cước vận chuyển ô tô: 0 đ / đơn vị",
            Location = new Point(18, 42),
            AutoSize = true,
            Font = new Font("Segoe UI", 14f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };

        pnlKetQua.Controls.Add(lblQuyDoi);
        pnlKetQua.Controls.Add(lblKetQua);
        pnlContent.Controls.Add(pnlKetQua);
    }

    private void LoadInitialData()
    {
        _suppressRecalc = true;

        txtDonGiaCaMay.Text = _donGiaCaMay.ToString("N0", ViVn);

        // 1. Kiểm tra cấu hình đã lưu trong kho bền vững
        var savedCfg = VanChuyenStorage.GetConfigOTo(_tenVatLieu);
        string? targetMa = _maDinhMucCu ?? savedCfg?.MaDinhMuc;
        if (savedCfg != null)
        {
            if (savedCfg.DonGiaCaMay > 0)
            {
                _donGiaCaMay = savedCfg.DonGiaCaMay;
                txtDonGiaCaMay.Text = _donGiaCaMay.ToString("N0", ViVn);
            }
        }

        // 2. Nhận diện định mức
        DinhMucVCOToItem? targetDm = null;
        if (!string.IsNullOrEmpty(targetMa))
        {
            targetDm = DinhMucVanChuyenDatabase.DanhSachOTo.FirstOrDefault(x => x.MaHieu == targetMa);
        }
        if (targetDm == null)
        {
            targetDm = DinhMucVanChuyenDatabase.NhanDienOTo(_tenVatLieu) ?? DinhMucVanChuyenDatabase.DanhSachOTo.First();
        }

        if (_layGiaMayFunc != null)
        {
            decimal g = _layGiaMayFunc(targetDm.MaMay, targetDm.TenMay);
            if (g > 0)
            {
                _donGiaCaMay = g;
                txtDonGiaCaMay.Text = _donGiaCaMay.ToString("N0", ViVn);
            }
        }

        // Chọn nhóm vật liệu
        int idxVl = cbVatLieu.Items.IndexOf(targetDm.LoaiVatLieu);
        cbVatLieu.SelectedIndex = idxVl >= 0 ? idxVl : 0;

        // Chọn loại xe
        int idxXe = -1;
        for (int i = 0; i < cbLoaiXe.Items.Count; i++)
        {
            if (cbLoaiXe.Items[i] is DinhMucVCOToItem item && item.MaHieu == targetDm.MaHieu)
            {
                idxXe = i;
                break;
            }
        }
        cbLoaiXe.SelectedIndex = idxXe >= 0 ? idxXe : 0;

        // 3. Nạp danh sách cung đoạn
        dgvCungDuong.Rows.Clear();
        LoadTemplatesToCombo();
        if (savedCfg != null && savedCfg.TongCuLyKm > 0 && savedCfg.CungDuongs != null && savedCfg.CungDuongs.Count > 0)
        {
            TongCuLyKm = savedCfg.TongCuLyKm;
            txtTongCuLy.Text = TongCuLyKm.ToString("0.###", ViVn);
            foreach (var cd in savedCfg.CungDuongs)
            {
                ThemCungDuongMoi(cd.DiemDau, cd.DiemCuoi, cd.TenDoanDuong, cd.CuLyKm, cd.LoaiDuong, insertIndex: -1);
            }
        }
        else
        {
            // Mở lần đầu: DỮ LIỆU HOÀN TOÀN TRỐNG
            txtTongCuLy.Text = "";
            TongCuLyKm = 0;
            dgvCungDuong.Rows.Clear();
        }

        _suppressRecalc = false;
        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    private void OnNhomVatLieuChanged()
    {
        if (cbVatLieu.SelectedItem is not string nhom) return;

        cbLoaiXe.Items.Clear();
        var xes = DinhMucVanChuyenDatabase.DanhSachOTo.Where(x => x.LoaiVatLieu == nhom).ToList();
        foreach (var x in xes)
        {
            cbLoaiXe.Items.Add(x);
        }

        if (cbLoaiXe.Items.Count > 0)
        {
            cbLoaiXe.SelectedIndex = 0;
        }
    }

    private void OnLoaiXeChanged()
    {
        if (cbLoaiXe.SelectedItem is not DinhMucVCOToItem dm) return;
        SelectedDinhMuc = dm;
        MaMay = dm.MaMay;

        if (_layGiaMayFunc != null)
        {
            decimal giaMoi = _layGiaMayFunc(dm.MaMay, dm.TenMay);
            if (giaMoi > 0)
            {
                DonGiaCaMay = giaMoi;
                txtDonGiaCaMay.Text = DonGiaCaMay.ToString("N0", ViVn);
            }
        }

        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    /// <summary>
    /// Khi người dùng nhập Tổng cự ly: Nếu bảng chưa có đoạn nào, tạo 1 đoạn mặc định toàn tuyến.
    /// </summary>
    private void OnTongCuLyUserSubmitted()
    {
        decimal ly = ParseVnFlex(txtTongCuLy.Text);
        if (ly <= 0)
        {
            TongCuLyKm = 0;
            txtTongCuLy.Text = "";
            dgvCungDuong.Rows.Clear();
            CapNhatThongTinDinhMucTheoTongCuLy();
            TinhToan();
            return;
        }

        TongCuLyKm = ly;
        txtTongCuLy.Text = TongCuLyKm.ToString("0.###", ViVn);

        // Nếu bảng đang trống hoặc chỉ có 1 dòng, cập nhật đoạn toàn tuyến
        if (dgvCungDuong.Rows.Count == 0)
        {
            _isUpdatingGridFromCode = true;
            ThemCungDuongMoi("Mỏ / Nguồn cung ứng", "Chân công trình", "Tuyến vận chuyển chính", TongCuLyKm, 3, insertIndex: -1);
            _isUpdatingGridFromCode = false;
        }
        else if (dgvCungDuong.Rows.Count == 1)
        {
            _isUpdatingGridFromCode = true;
            dgvCungDuong["CuLy", 0].Value = TongCuLyKm.ToString("0.###", ViVn);
            _isUpdatingGridFromCode = false;
        }

        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    private void ChenDoanVaoViTriDangChon()
    {
        int insertIdx = dgvCungDuong.CurrentRow != null ? dgvCungDuong.CurrentRow.Index + 1 : dgvCungDuong.Rows.Count;
        string dDau = "";
        if (dgvCungDuong.CurrentRow != null)
        {
            dDau = dgvCungDuong.CurrentRow.Cells["DiemCuoi"].Value?.ToString() ?? "";
        }
        ThemCungDuongMoi(dDau, "", $"Đoạn {insertIdx + 1}", 0, 3, insertIndex: insertIdx);
    }

    private void CapNhatTongCuLyTuBang()
    {
        decimal tong = 0;
        foreach (DataGridViewRow r in dgvCungDuong.Rows)
        {
            tong += ParseVnFlex(r.Cells["CuLy"].Value?.ToString());
        }

        TongCuLyKm = tong;
        txtTongCuLy.Text = TongCuLyKm > 0 ? TongCuLyKm.ToString("0.###", ViVn) : "";
        CapNhatThongTinDinhMucTheoTongCuLy();
    }

    private void CapNhatThongTinDinhMucTheoTongCuLy()
    {
        if (SelectedDinhMuc == null) return;
        var dm = SelectedDinhMuc;

        if (TongCuLyKm <= 0)
        {
            lblDmInfo.Text = "Nhập các đoạn theo hành trình thực tế từ nguồn cung đến công trình để phần mềm tự động phân bổ";
            return;
        }

        var parts = new List<string>();
        parts.Add($"Đm1 (<=1km) = {dm.Dm1.ToString("0.###", ViVn)}");
        if (TongCuLyKm > 1) parts.Add($"Đm2 (1-10km) = {dm.Dm2.ToString("0.###", ViVn)}");
        if (TongCuLyKm > 10) parts.Add($"Đm3 (10-60km) = {dm.Dm3.ToString("0.###", ViVn)}");
        if (TongCuLyKm > 60) parts.Add($"Đm4 (>60km) = {(dm.Dm3 * 0.95m).ToString("0.###", ViVn)}");

        lblDmInfo.Text = $"Các định mức áp dụng theo cự ly {TongCuLyKm.ToString("0.###", ViVn)} km ({dm.DonViDinhMuc}/1km): " + string.Join("  |  ", parts);
    }

    private void ThemCungDuongMoi(string diemDau, string diemCuoi, string tenDoan, decimal cuLy, int loaiDuong, int insertIndex = -1)
    {
        int loaiIdx = Math.Max(0, Math.Min(5, loaiDuong - 1));
        var loaiStr = dgvCungDuong.Columns["LoaiDuong"] is DataGridViewComboBoxColumn cbc ? cbc.Items[loaiIdx].ToString() : "Loại 3 (k=1,00)";
        decimal heSoK = new CungDuongVanChuyen { LoaiDuong = loaiDuong }.HeSoK;

        object[] rowData = new object[]
        {
            "1",
            diemDau,
            diemCuoi,
            tenDoan, 
            cuLy > 0 ? cuLy.ToString("0.###", ViVn) : "0", 
            loaiStr, 
            heSoK.ToString("0.##", ViVn), 
            ""
        };

        if (insertIndex >= 0 && insertIndex <= dgvCungDuong.Rows.Count)
        {
            dgvCungDuong.Rows.Insert(insertIndex, rowData);
        }
        else
        {
            dgvCungDuong.Rows.Add(rowData);
        }

        // Đánh lại STT liên tục
        for (int i = 0; i < dgvCungDuong.Rows.Count; i++)
        {
            dgvCungDuong["STT", i].Value = (i + 1).ToString();
        }

        if (!_isUpdatingGridFromCode)
        {
            CapNhatTongCuLyTuBang();
            TinhToan();
        }
    }

    private void XoaCungDuongChon()
    {
        if (dgvCungDuong.CurrentRow != null)
        {
            dgvCungDuong.Rows.Remove(dgvCungDuong.CurrentRow);
            for (int i = 0; i < dgvCungDuong.Rows.Count; i++)
            {
                dgvCungDuong["STT", i].Value = (i + 1).ToString();
            }

            CapNhatTongCuLyTuBang();
            TinhToan();
        }
    }

    private void DgvCungDuong_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        var colName = dgvCungDuong.Columns[e.ColumnIndex].Name;

        if (colName == "LoaiDuong")
        {
            var val = dgvCungDuong["LoaiDuong", e.RowIndex].Value?.ToString() ?? "";
            int loai = ParseLoaiDuong(val);

            decimal k = new CungDuongVanChuyen { LoaiDuong = loai }.HeSoK;
            dgvCungDuong["HeSoK", e.RowIndex].Value = k.ToString("0.##", ViVn);
        }
    }

    private static int ParseLoaiDuong(string val)
    {
        if (string.IsNullOrWhiteSpace(val)) return 3;
        var trim = val.Trim();
        if (trim.StartsWith("Loại 1")) return 1;
        if (trim.StartsWith("Loại 2")) return 2;
        if (trim.StartsWith("Loại 3")) return 3;
        if (trim.StartsWith("Loại 4")) return 4;
        if (trim.StartsWith("Loại 5")) return 5;
        if (trim.StartsWith("Loại 6")) return 6;
        return 3;
    }

    private void ReadCungDuongsFromGrid()
    {
        CungDuongs.Clear();
        foreach (DataGridViewRow row in dgvCungDuong.Rows)
        {
            if (row.IsNewRow) continue;
            var dDau = row.Cells["DiemDau"].Value?.ToString() ?? "";
            var dCuoi = row.Cells["DiemCuoi"].Value?.ToString() ?? "";
            var ten = row.Cells["TenDoan"].Value?.ToString() ?? "";
            decimal cuLy = ParseVnFlex(row.Cells["CuLy"].Value?.ToString());
            var loaiStr = row.Cells["LoaiDuong"].Value?.ToString() ?? "";
            int loai = ParseLoaiDuong(loaiStr);

            CungDuongs.Add(new CungDuongVanChuyen
            {
                DiemDau = dDau,
                DiemCuoi = dCuoi,
                TenDoanDuong = ten,
                CuLyKm = cuLy,
                LoaiDuong = loai
            });
        }
    }

    private void LoadTemplatesToCombo()
    {
        if (cbTemplates == null) return;
        cbTemplates.Items.Clear();
        cbTemplates.Items.Add("-- Chọn mẫu tuyến đường đã lưu --");
        var templates = VanChuyenStorage.GetAllTemplates();
        foreach (var t in templates)
        {
            cbTemplates.Items.Add(t);
        }
        cbTemplates.SelectedIndex = 0;
    }

    private void ApDungTemplate(TuyenDuongTemplate t)
    {
        _isUpdatingGridFromCode = true;
        try
        {
            dgvCungDuong.Rows.Clear();
            foreach (var cd in t.CungDuongs)
            {
                ThemCungDuongMoi(cd.DiemDau, cd.DiemCuoi, cd.TenDoanDuong, cd.CuLyKm, cd.LoaiDuong, insertIndex: -1);
            }
            CapNhatTongCuLyTuBang();
            TinhToan();
        }
        finally
        {
            _isUpdatingGridFromCode = false;
        }
    }

    private void OnLuuTemplateClicked()
    {
        ReadCungDuongsFromGrid();
        if (CungDuongs.Count == 0 || CungDuongs.All(c => c.CuLyKm <= 0))
        {
            MessageBox.Show("Vui lòng nhập ít nhất một đoạn đường có cự ly trước khi lưu mẫu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string defaultName = "";
        var first = CungDuongs.FirstOrDefault(c => !string.IsNullOrWhiteSpace(c.DiemDau));
        var last = CungDuongs.LastOrDefault(c => !string.IsNullOrWhiteSpace(c.DiemCuoi));
        if (first != null && last != null)
        {
            defaultName = $"{first.DiemDau} → {last.DiemCuoi}";
        }
        else
        {
            defaultName = $"Tuyến {TongCuLyKm:0.###}km ({_tenVatLieu})";
        }

        using var inputForm = new Form
        {
            Text = "Lưu mẫu tuyến đường",
            Size = new Size(480, 185),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            Font = new Font("Segoe UI", 9.5f)
        };
        var lblPrompt = new Label { Text = "Nhập tên gợi nhớ cho mẫu tuyến đường vận chuyển:", Location = new Point(20, 15), AutoSize = true };
        var txtName = new TextBox { Text = defaultName, Location = new Point(20, 44), Width = 420, Font = new Font("Segoe UI", 10f) };
        var btnOk = new Button { Text = "✔ Lưu mẫu", DialogResult = DialogResult.OK, Location = new Point(245, 90), Width = 100, Height = 34, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        var btnCancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Location = new Point(355, 90), Width = 85, Height = 34 };
        inputForm.Controls.AddRange(new Control[] { lblPrompt, txtName, btnOk, btnCancel });
        inputForm.AcceptButton = btnOk;
        inputForm.CancelButton = btnCancel;

        if (inputForm.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(txtName.Text))
        {
            var tenMoi = txtName.Text.Trim();
            VanChuyenStorage.SaveTemplate(tenMoi, CungDuongs);
            LoadTemplatesToCombo();
            for (int i = 0; i < cbTemplates.Items.Count; i++)
            {
                if (cbTemplates.Items[i] is TuyenDuongTemplate temp && temp.TenTemplate.Equals(tenMoi, StringComparison.OrdinalIgnoreCase))
                {
                    cbTemplates.SelectedIndex = i;
                    break;
                }
            }
            MessageBox.Show($"Đã lưu mẫu tuyến đường \"{tenMoi}\" thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void OnXoaTemplateClicked()
    {
        if (cbTemplates.SelectedItem is TuyenDuongTemplate t)
        {
            var confirm = MessageBox.Show(
                $"Bạn có chắc chắn muốn xóa mẫu tuyến đường \"{t.TenTemplate}\"?",
                "Xác nhận xóa mẫu tuyến",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm == DialogResult.Yes)
            {
                VanChuyenStorage.DeleteTemplate(t.TenTemplate);
                LoadTemplatesToCombo();
            }
        }
        else
        {
            MessageBox.Show("Vui lòng chọn một mẫu tuyến đường trong danh sách để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private class PhanDoanNac
    {
        public decimal CuLyThuc { get; set; }
        public int LoaiDuong { get; set; }
        public decimal HeSoK { get; set; }
    }

    /// <summary>
    /// Thuật toán phân rã cự ly tích lũy đa cung đoạn chuẩn xác 100% theo hướng dẫn và ví dụ tính toán
    /// tại Chương XII - Thông tư số 38/2026/TT-BXD (trang 686).
    /// </summary>
    private void TinhToan()
    {
        if (_suppressRecalc) return;
        if (cbLoaiXe.SelectedItem is not DinhMucVCOToItem dm) return;

        ReadCungDuongsFromGrid();
        DonGiaCaMay = ParseVnFlex(txtDonGiaCaMay.Text);

        decimal tongCuLy = CungDuongs.Sum(x => x.CuLyKm);
        TongCuLyKm = tongCuLy;

        if (tongCuLy <= 0 || CungDuongs.Count == 0)
        {
            lblHuongDan.Visible = true;
            lblNac1.Visible = false;
            lblNac2.Visible = false;
            lblNac3.Visible = false;
            lblNac4.Visible = false;
            lblTongCaXe.Text = "Tổng hao phí ca máy = 0 ca";
            KetQuaCuocOTo = 0;
            lblQuyDoi.Text = "Chưa có cự ly vận chuyển";
            lblKetQua.Text = $"Cước vận chuyển ô tô: 0 đ / {_donViVatLieu}";
            pnlChietTinh.Height = 110;
            return;
        }

        lblHuongDan.Visible = false;

        // PHÂN RÃ THEO HÀNH TRÌNH TÍCH LŨY (Chuẩn Thông tư 38/2026/TT-BXD)
        var nac1Segments = new List<PhanDoanNac>();
        var nac2Segments = new List<PhanDoanNac>();
        var nac3Segments = new List<PhanDoanNac>();
        var nac4Segments = new List<PhanDoanNac>();

        decimal accumulated = 0;

        foreach (var cd in CungDuongs)
        {
            if (cd.CuLyKm <= 0) continue;

            decimal start = accumulated;
            decimal end = accumulated + cd.CuLyKm;
            accumulated = end;

            // Nấc 1: [0, 1]
            decimal o1 = Math.Max(0, Math.Min(end, 1m) - Math.Max(start, 0m));
            if (o1 > 0)
            {
                nac1Segments.Add(new PhanDoanNac { CuLyThuc = o1, LoaiDuong = cd.LoaiDuong, HeSoK = cd.HeSoK });
            }

            // Nấc 2: [1, 10]
            decimal o2 = Math.Max(0, Math.Min(end, 10m) - Math.Max(start, 1m));
            if (o2 > 0)
            {
                nac2Segments.Add(new PhanDoanNac { CuLyThuc = o2, LoaiDuong = cd.LoaiDuong, HeSoK = cd.HeSoK });
            }

            // Nấc 3: [10, 60]
            decimal o3 = Math.Max(0, Math.Min(end, 60m) - Math.Max(start, 10m));
            if (o3 > 0)
            {
                nac3Segments.Add(new PhanDoanNac { CuLyThuc = o3, LoaiDuong = cd.LoaiDuong, HeSoK = cd.HeSoK });
            }

            // Nấc 4: > 60
            decimal o4 = Math.Max(0, end - Math.Max(start, 60m));
            if (o4 > 0)
            {
                nac4Segments.Add(new PhanDoanNac { CuLyThuc = o4, LoaiDuong = cd.LoaiDuong, HeSoK = cd.HeSoK });
            }
        }

        // Tính toán chiết tính từng nấc
        decimal l1Thuc = nac1Segments.Sum(x => x.CuLyThuc);
        decimal q1 = nac1Segments.Sum(x => x.CuLyThuc * x.HeSoK);
        decimal caNac1 = dm.Dm1 * q1;

        decimal l2Thuc = nac2Segments.Sum(x => x.CuLyThuc);
        decimal q2 = nac2Segments.Sum(x => x.CuLyThuc * x.HeSoK);
        decimal caNac2 = dm.Dm2 * q2;

        decimal l3Thuc = nac3Segments.Sum(x => x.CuLyThuc);
        decimal q3 = nac3Segments.Sum(x => x.CuLyThuc * x.HeSoK);
        decimal caNac3 = dm.Dm3 * q3;

        decimal l4Thuc = nac4Segments.Sum(x => x.CuLyThuc);
        decimal q4 = nac4Segments.Sum(x => x.CuLyThuc * x.HeSoK);
        decimal caNac4 = (dm.Dm3 * 0.95m) * q4;

        decimal tongCaXe = caNac1 + caNac2 + caNac3 + caNac4;

        // DIỄN GIẢI ĐÚNG VÍ DỤ NGUYÊN BẢN CỦA THÔNG TƯ 38 TRANG 686 (DÙNG DẤU PHẨY THẬP PHÂN CHUẨN VN)
        int curY = 40;

        string FormatDienGiaiNac(List<PhanDoanNac> segs)
        {
            var parts = segs.Select(s => $"{s.CuLyThuc.ToString("0.###", ViVn)} x k{s.LoaiDuong} ({s.HeSoK.ToString("0.##", ViVn)})").ToList();
            return string.Join(" + ", parts);
        }

        if (l1Thuc > 0)
        {
            lblNac1.Visible = true;
            lblNac1.Location = new Point(20, curY);
            lblNac1.Text = $"• Nấc 1 (Phạm vi <= 1km):  L = {l1Thuc.ToString("0.###", ViVn)} km  ==>  Đm1 x ({FormatDienGiaiNac(nac1Segments)}) = {dm.Dm1.ToString("0.###", ViVn)} x {q1.ToString("0.###", ViVn)} = {caNac1.ToString("0.####", ViVn)} ca";
            curY += 26;
        }
        else
        {
            lblNac1.Visible = false;
        }

        if (l2Thuc > 0)
        {
            lblNac2.Visible = true;
            lblNac2.Location = new Point(20, curY);
            lblNac2.Text = $"• Nấc 2 (Cự ly 1-10km):   L = {l2Thuc.ToString("0.###", ViVn)} km  ==>  Đm2 x ({FormatDienGiaiNac(nac2Segments)}) = {dm.Dm2.ToString("0.###", ViVn)} x {q2.ToString("0.###", ViVn)} = {caNac2.ToString("0.####", ViVn)} ca";
            curY += 26;
        }
        else
        {
            lblNac2.Visible = false;
        }

        if (l3Thuc > 0)
        {
            lblNac3.Visible = true;
            lblNac3.Location = new Point(20, curY);
            lblNac3.Text = $"• Nấc 3 (Cự ly 10-60km):  L = {l3Thuc.ToString("0.###", ViVn)} km  ==>  Đm3 x ({FormatDienGiaiNac(nac3Segments)}) = {dm.Dm3.ToString("0.###", ViVn)} x {q3.ToString("0.###", ViVn)} = {caNac3.ToString("0.####", ViVn)} ca";
            curY += 26;
        }
        else
        {
            lblNac3.Visible = false;
        }

        if (l4Thuc > 0)
        {
            lblNac4.Visible = true;
            lblNac4.Location = new Point(20, curY);
            lblNac4.Text = $"• Nấc 4 (Cự ly > 60km):   L = {l4Thuc.ToString("0.###", ViVn)} km  ==>  (Đm3 x 0,95) x ({FormatDienGiaiNac(nac4Segments)}) = {(dm.Dm3 * 0.95m).ToString("0.###", ViVn)} x {q4.ToString("0.###", ViVn)} = {caNac4.ToString("0.####", ViVn)} ca";
            curY += 26;
        }
        else
        {
            lblNac4.Visible = false;
        }

        lblTongCaXe.Location = new Point(20, curY + 6);
        lblTongCaXe.Text = $"Tổng cự ly: {tongCuLy.ToString("0.###", ViVn)} km  ==>  Tổng hao phí ca máy: {tongCaXe.ToString("0.####", ViVn)} ca (cho {dm.DonViDinhMuc})";
        pnlChietTinh.Height = curY + 42;

        decimal gia1Dm = (tongCaXe * DonGiaCaMay) / 10m;

        decimal heSoQuyDoi = DinhMucVanChuyenDatabase.TinhHeSoQuyDoiOTo(_donViVatLieu, dm.DonViDinhMuc);
        KetQuaCuocOTo = Math.Round(gia1Dm * heSoQuyDoi * 10m, 0, MidpointRounding.AwayFromZero);

        string dvGoc = dm.DonViDinhMuc.Replace("10", "");
        if (heSoQuyDoi * 10m != 1.0m)
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dvGoc} = {gia1Dm.ToString("N0", ViVn)} đ   (Quy đổi về 1 {_donViVatLieu} theo đơn giá)";
        }
        else
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dvGoc} = {gia1Dm.ToString("N0", ViVn)} đ   (Cùng đơn vị tính với vật tư)";
        }

        lblKetQua.Text = $"Cước vận chuyển ô tô: {KetQuaCuocOTo.ToString("N0", ViVn)} đ / {_donViVatLieu}";
    }

    private void MoPopupTraCuuDuong()
    {
        using var frm = new Form
        {
            Text = "Tra cứu phân loại đường bộ TP Đà Nẵng (QĐ 2035/QĐ-UBND)",
            Size = new Size(950, 560),
            StartPosition = FormStartPosition.CenterParent,
            Font = new Font("Segoe UI", 9.5f)
        };

        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            BackgroundColor = Color.White
        };

        dgv.Columns.Add("TuyenDuong", "Tuyến đường");
        dgv.Columns.Add("DiaPhan", "Địa phận");
        dgv.Columns.Add("Doan", "Từ Km đến Km");
        dgv.Columns.Add("ChieuDai", "Chiều dài (km)");
        dgv.Columns.Add("LoaiDuong", "Loại đường");
        dgv.Columns.Add("GhiChu", "Ghi chú");

        dgv.Columns["ChieuDai"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        dgv.Columns["LoaiDuong"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.Columns["GhiChu"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

        foreach (var item in DinhMucVanChuyenDatabase.DanhSachDuongDaNang)
        {
            dgv.Rows.Add(item.TuyenDuong, item.DiaPhan, item.TuKmDenKm, item.ChieuDaiKm.ToString("0.###", ViVn), $"Loại {item.LoaiDuong}", item.GhiChu);
        }

        var pnlBot = new Panel { Dock = DockStyle.Bottom, Height = 55, BackColor = Color.FromArgb(245, 245, 245) };
        var btnChon = new Button
        {
            Text = "Chọn tuyến đường này",
            DialogResult = DialogResult.OK,
            Width = 190,
            Height = 38,
            Location = new Point(pnlBot.Width - 210, 8),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(0, 123, 255),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
        };
        btnChon.FlatAppearance.BorderSize = 0;
        btnChon.Click += (s, e) =>
        {
            if (dgv.CurrentRow != null)
            {
                int idx = dgv.CurrentRow.Index;
                if (idx >= 0 && idx < DinhMucVanChuyenDatabase.DanhSachDuongDaNang.Count)
                {
                    var d = DinhMucVanChuyenDatabase.DanhSachDuongDaNang[idx];
                    ThemCungDuongMoi(d.DiaPhan, d.GhiChu, $"{d.TuyenDuong} ({d.TuKmDenKm})", d.ChieuDaiKm, d.LoaiDuong, insertIndex: -1);
                }
            }
            frm.Close();
        };

        pnlBot.Controls.Add(btnChon);
        frm.Controls.Add(dgv);
        frm.Controls.Add(pnlBot);

        dgv.CellDoubleClick += (s, e) => btnChon.PerformClick();
        frm.ShowDialog();
    }

    /// <summary>
    /// Hàm parse số linh hoạt hỗ trợ người dùng nhập dấu chấm hoặc dấu phẩy cho số thập phân,
    /// phân cách hàng nghìn chuẩn Việt Nam.
    /// </summary>
    public static decimal ParseVnFlex(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        string s = text.Trim();

        // Nếu chuỗi chứa cả dấu chấm và phẩy (ví dụ: 1.245,55 hoặc 1,245.55)
        if (s.Contains(".") && s.Contains(","))
        {
            int lastDot = s.LastIndexOf('.');
            int lastComma = s.LastIndexOf(',');
            if (lastComma > lastDot)
            {
                // Chuẩn Việt Nam: 1.245,55 -> bỏ chấm, đổi phẩy thành chấm
                s = s.Replace(".", "").Replace(",", ".");
            }
            else
            {
                // Chuẩn US: 1,245.55 -> bỏ phẩy
                s = s.Replace(",", "");
            }
        }
        else if (s.Contains(","))
        {
            // Chỉ chứa dấu phẩy: người dùng gõ số thập phân (ví dụ: 1245,55 hoặc 0,3)
            s = s.Replace(",", ".");
        }
        else if (s.Contains("."))
        {
            // Chỉ chứa dấu chấm:
            int countDot = s.Count(c => c == '.');
            if (countDot > 1)
            {
                // Nhiều dấu chấm (ví dụ: 1.040.324) -> dấu phân cách hàng nghìn
                s = s.Replace(".", "");
            }
            else
            {
                // Có 1 dấu chấm: người dùng gõ kiểu thập phân (ví dụ: 1245.55 hoặc 0.3)
                // Giữ nguyên dấu chấm để parse InvariantCulture
            }
        }

        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val) ? val : 0;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
