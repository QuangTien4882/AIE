using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.ExcelAddIn.Services;

namespace AIE.ExcelAddIn.Forms;

public class TinhCuocVCBoForm : Form
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
    private decimal _donGiaNhanCong;
    private readonly string? _maDinhMucCu;

    public decimal KetQuaCuocBo { get; private set; }
    public DinhMucVCBoItem? SelectedDinhMuc { get; private set; }
    public decimal DonGiaNhanCong { get; private set; }
    public decimal TongCuLyMet { get; private set; } = 0;
    public List<DoanVanChuyenBo> DoanBos { get; private set; } = new();

    // UI Controls
    private Label lblCuLy;
    private TextBox txtTongCuLy;
    private Label lblGiaNC;
    private TextBox txtDonGiaNC;
    private Label lblLoaiDM;
    private ComboBox cbDinhMuc;
    private Label lblDvtDinhMuc;
    private Label lblDmInfo;

    // Cao tầng Controls (theo Thông tư số 38/2026/TT-BXD)
    private CheckBox chkCaoTang;
    private Label lblTang;
    private NumericUpDown numTang;
    private CheckBox chkThangMay;
    private Label lblHeSoTangInfo;

    private DataGridView dgvDoanBo;
    private Button btnThemDoan;
    private Button btnXoaDoan;

    private Panel pnlChietTinh;
    private Label lblHuongDan;
    private Label lblNac1;
    private Label lblNac2;
    private Label lblTongCong;

    private Label lblQuyDoi;
    private Label lblKetQua;

    private string? _maNhanCong;
    private bool _suppressRecalc = false;
    private bool _isUpdatingGridFromCode = false;

    public TinhCuocVCBoForm(
        string tenVatLieu, 
        string donViVatLieu, 
        decimal donGiaNhanCong = 0, 
        decimal giaTriHienTai = 0,
        string? maDinhMucCu = null,
        string? maNhanCong = null)
    {
        _tenVatLieu = tenVatLieu;
        _donViVatLieu = string.IsNullOrWhiteSpace(donViVatLieu) ? "ĐVT" : donViVatLieu.Trim();
        _donGiaNhanCong = donGiaNhanCong > 0 ? donGiaNhanCong : 254498m;
        _maDinhMucCu = maDinhMucCu;
        _maNhanCong = maNhanCong;

        InitializeComponent();
        LoadInitialData();

        if (giaTriHienTai > 0 && KetQuaCuocBo == 0)
        {
            KetQuaCuocBo = giaTriHienTai;
        }
    }

    private void InitializeComponent()
    {
        this.Text = $"Tính chi phí vận chuyển bộ (thủ công) - {_tenVatLieu}";

        var area = Screen.PrimaryScreen.WorkingArea;
        int targetW = Math.Min(1200, area.Width - 40);
        int targetH = Math.Min(880, area.Height - 40);
        this.ClientSize = new Size(targetW, targetH);
        this.MinimumSize = new Size(1020, 720);
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
            Text = "TÍNH CHI PHÍ VẬN CHUYỂN BỘ (THỦ CÔNG)",
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
            ReadDoanBosFromGrid();
            VanChuyenStorage.SaveConfigBo(
                _tenVatLieu, 
                SelectedDinhMuc?.MaHieu ?? "", 
                TongCuLyMet, 
                1.0m, 
                (int)numTang.Value, 
                KetQuaCuocBo,
                DonGiaNhanCong,
                _maNhanCong ?? "");
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
        // CARD 1: THÔNG TIN CHUNG & ĐỊNH MỨC (BỐ CỤC 3 HÀNG RỘNG RÃI, KHÔNG BAO GIỜ TRÀN)
        // =========================================================================
        var cardDinhMuc = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 185),
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
            Text = "📋  THÔNG TIN CHUNG & ĐỊNH MỨC VẬN CHUYỂN BỘ (AM.21000 - Thông tư số 38/2026/TT-BXD)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 9),
            BackColor = Color.Transparent
        };
        cardDinhMuc.Controls.Add(lblCard1Title);

        // HÀNG 1: TỔNG CỰ LY & GIÁ NHÂN CÔNG
        lblCuLy = new Label { Text = "Tổng cự ly vận chuyển bộ (m):", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
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

        lblGiaNC = new Label { Text = "Giá NC nhóm I (đ):", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular) };
        txtDonGiaNC = new TextBox
        {
            Size = new Size(140, 26),
            TextAlign = HorizontalAlignment.Right,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold)
        };
        txtDonGiaNC.TextChanged += (s, e) => TinhToan();

        // HÀNG 2: LOẠI CÔNG TÁC ĐỊNH MỨC (RỘNG RÃI TOÀN HÀNG)
        lblLoaiDM = new Label { Text = "Loại công tác định mức:", AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular) };
        cbDinhMuc = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Height = 26,
            DropDownWidth = 700,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };
        foreach (var item in DinhMucVanChuyenDatabase.DanhSachBo)
        {
            cbDinhMuc.Items.Add(item);
        }
        cbDinhMuc.SelectedIndexChanged += (s, e) => OnDinhMucChanged();

        // HÀNG 3: ĐƠN VỊ TÍNH & QUY ĐỊNH CAO TẦNG
        lblDvtDinhMuc = new Label
        {
            Text = "ĐVT: m³",
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 102, 204)
        };

        chkCaoTang = new CheckBox
        {
            Text = "Vận chuyển lên cao tầng (tầng 2 trở lên)",
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(180, 83, 9)
        };
        chkCaoTang.CheckedChanged += (s, e) => OnCaoTangSettingsChanged();

        lblTang = new Label { Text = "Lên tầng:", AutoSize = true, Visible = false };
        numTang = new NumericUpDown
        {
            Size = new Size(60, 26),
            Minimum = 1,
            Maximum = 50,
            Value = 2,
            Visible = false
        };
        numTang.ValueChanged += (s, e) => OnCaoTangSettingsChanged();

        chkThangMay = new CheckBox
        {
            Text = "Dùng thang máy / vận thăng / cần cẩu tháp (k=1,00)",
            AutoSize = true,
            Visible = false
        };
        chkThangMay.CheckedChanged += (s, e) => OnCaoTangSettingsChanged();

        lblHeSoTangInfo = new Label
        {
            Text = "",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Italic),
            ForeColor = Color.FromArgb(40, 167, 69)
        };

        // HÀNG 4: THÔNG TIN ĐỊNH MỨC TÓM TẮT
        lblDmInfo = new Label
        {
            Text = "Nhập Tổng cự ly vận chuyển bộ (m) ở trên hoặc thêm các đoạn ở dưới để áp dụng định mức",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        cardDinhMuc.Controls.AddRange(new Control[] {
            lblCuLy, txtTongCuLy, lblGiaNC, txtDonGiaNC,
            lblLoaiDM, cbDinhMuc,
            lblDvtDinhMuc, chkCaoTang, lblTang, numTang, chkThangMay, lblHeSoTangInfo,
            lblDmInfo
        });

        Action layoutCard1 = () =>
        {
            // Hàng 1 (Y = 46): Cự ly và Giá nhân công
            lblCuLy.Location = new Point(16, 49);
            txtTongCuLy.Location = new Point(lblCuLy.Right + 10, 45);

            lblGiaNC.Location = new Point(txtTongCuLy.Right + 35, 49);
            txtDonGiaNC.Location = new Point(lblGiaNC.Right + 10, 45);

            // Hàng 2 (Y = 82): Loại công tác co giãn tự động theo toàn bộ bề ngang Card
            lblLoaiDM.Location = new Point(16, 84);
            int cbLeft = lblLoaiDM.Right + 10;
            int cbWidth = Math.Max(350, cardDinhMuc.ClientSize.Width - cbLeft - 25);
            cbDinhMuc.Location = new Point(cbLeft, 80);
            cbDinhMuc.Size = new Size(cbWidth, 26);

            // Hàng 3 (Y = 118): ĐVT và Tùy chọn cao tầng
            lblDvtDinhMuc.Location = new Point(16, 120);
            chkCaoTang.Location = new Point(lblDvtDinhMuc.Right + 25, 119);
            lblTang.Location = new Point(chkCaoTang.Right + 15, 120);
            numTang.Location = new Point(lblTang.Right + 8, 117);
            chkThangMay.Location = new Point(numTang.Right + 15, 119);

            lblHeSoTangInfo.Location = new Point(chkThangMay.Right + 15, 120);

            // Hàng 4 (Y = 152): Hướng dẫn định mức
            lblDmInfo.Location = new Point(16, 154);
        };
        cardDinhMuc.Layout += (s, e) => layoutCard1();
        cardDinhMuc.Resize += (s, e) => layoutCard1();
        layoutCard1();

        pnlContent.Controls.Add(cardDinhMuc);
        y += 200;

        // =========================================================================
        // CARD 2: BẢNG CÁC ĐOẠN VẬN CHUYỂN BỘ (CARD PANEL HIỆN ĐẠI)
        // =========================================================================
        var cardCungDuong = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 275),
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
            Text = "🚶  CHI TIẾT CÁC ĐOẠN VẬN CHUYỂN BỘ (Tự động tính toán theo cự ly và hệ số cao tầng)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 9),
            BackColor = Color.Transparent
        };
        cardCungDuong.Controls.Add(lblCard2Title);

        dgvDoanBo = new DataGridView
        {
            Location = new Point(16, 45),
            Size = new Size(cardCungDuong.ClientSize.Width - 32, 170),
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
        dgvDoanBo.RowTemplate.Height = 29;

        // Click 1 lần là edit được ngay, nếu là dropdown thì bung danh sách ngay lập tức
        dgvDoanBo.CellClick += (s, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                dgvDoanBo.BeginEdit(true);
                if (dgvDoanBo.EditingControl is ComboBox cb)
                {
                    cb.DroppedDown = true;
                }
            }
        };

        // Bắt phím Enter trong DataGridView để không nhảy ra sheet Excel
        dgvDoanBo.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                dgvDoanBo.EndEdit();
                e.Handled = true;
            }
        };

        var colStt = new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 45, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
        var colNac = new DataGridViewComboBoxColumn { Name = "NacCuLy", HeaderText = "Nấc định mức", Width = 150 };
        colNac.Items.AddRange("10m khởi điểm", "Mỗi 10m tiếp theo");
        var colDiemDau = new DataGridViewTextBoxColumn { Name = "DiemDau", HeaderText = "Điểm đầu", Width = 140 };
        var colDiemCuoi = new DataGridViewTextBoxColumn { Name = "DiemCuoi", HeaderText = "Điểm cuối", Width = 140 };
        var colCuLy = new DataGridViewTextBoxColumn { Name = "CuLy", HeaderText = "Cự ly (m)", Width = 100, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
        var colDiaHinh = new DataGridViewComboBoxColumn { Name = "DiaHinh", HeaderText = "Địa hình / Độ dốc", Width = 185 };
        colDiaHinh.Items.AddRange(
            "Bằng phẳng, dốc <= 7° (k=1,00)",
            "Dốc <= 10° (k=1,20)",
            "Dốc <= 15° (k=1,35)",
            "Dốc <= 20° (k=1,70)",
            "Dốc <= 25° (k=2,00)",
            "Dốc <= 30° (k=2,50)",
            "Gồ ghề, lởm chởm (k=1,50)",
            "Trơn, lầy lún (k=2,50)");
        var colTang = new DataGridViewTextBoxColumn { Name = "SoTang", HeaderText = "Tầng", Width = 60, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
        var colHeSo = new DataGridViewTextBoxColumn { Name = "HeSoK", HeaderText = "Hệ số k", Width = 85, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) } };
        var colHaoPhi = new DataGridViewTextBoxColumn { Name = "HaoPhiCong", HeaderText = "Hao phí (công)", Width = 120, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) } };
        var colGhiChu = new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill };

        dgvDoanBo.Columns.AddRange(colStt, colNac, colDiemDau, colDiemCuoi, colCuLy, colDiaHinh, colTang, colHeSo, colHaoPhi, colGhiChu);
        dgvDoanBo.CellValueChanged += DgvDoanBo_CellValueChanged;
        dgvDoanBo.CellEndEdit += (s, e) => 
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var colName = dgvDoanBo.Columns[e.ColumnIndex].Name;
                if (colName == "CuLy")
                {
                    decimal val = TinhCuocVCOToForm.ParseVnFlex(dgvDoanBo["CuLy", e.RowIndex].Value?.ToString());
                    dgvDoanBo["CuLy", e.RowIndex].Value = val > 0 ? val.ToString("0.##", ViVn) : "0";
                }
            }
            if (!_isUpdatingGridFromCode)
            {
                CapNhatTongCuLyTuBang();
            }
            TinhToan();
        };

        dgvDoanBo.CellFormatting += (s, e) =>
        {
            if (e.RowIndex >= 0 && dgvDoanBo.Columns[e.ColumnIndex].Name == "CuLy" && e.Value != null)
            {
                decimal v = TinhCuocVCOToForm.ParseVnFlex(e.Value.ToString());
                e.Value = v > 0 ? v.ToString("0.##", ViVn) : "0";
                e.FormattingApplied = true;
            }
        };

        var pnlButtons = new FlowLayoutPanel
        {
            Location = new Point(16, 225),
            Size = new Size(cardCungDuong.ClientSize.Width - 32, 42),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            WrapContents = false,
            AutoScroll = false
        };

        btnThemDoan = new Button
        {
            Text = "+ Thêm đoạn vận chuyển",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(16, 0, 16, 0),
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
            string nacMacDinh = "Mỗi 10m tiếp theo";
            string diemDauMacDinh = "";
            int insertIdx = -1;
            if (dgvDoanBo.CurrentRow != null)
            {
                insertIdx = dgvDoanBo.CurrentRow.Index + 1;
                nacMacDinh = dgvDoanBo.CurrentRow.Cells["NacCuLy"].Value?.ToString() ?? "Mỗi 10m tiếp theo";
                diemDauMacDinh = dgvDoanBo.CurrentRow.Cells["DiemCuoi"].Value?.ToString() ?? "";
            }
            int tangHienTai = chkCaoTang.Checked ? (int)numTang.Value : 1;
            ThemDoanBoMoi(diemDauMacDinh, "", 0, 1.0m, tangHienTai, nacMacDinh, insertIdx);
        };

        btnXoaDoan = new Button
        {
            Text = "- Xóa đoạn chọn",
            Height = 36,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(16, 0, 16, 0),
            Margin = new Padding(0, 0, 0, 0),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter
        };
        btnXoaDoan.FlatAppearance.BorderColor = Color.FromArgb(206, 212, 218);
        btnXoaDoan.Click += (s, e) => XoaDoanBoChon();

        pnlButtons.Controls.Add(btnThemDoan);
        pnlButtons.Controls.Add(btnXoaDoan);

        cardCungDuong.Controls.Add(dgvDoanBo);
        cardCungDuong.Controls.Add(pnlButtons);
        pnlContent.Controls.Add(cardCungDuong);
        y += 290;

        // =========================================================================
        // CARD 3: TỔNG HỢP CHIẾT TÍNH VẬN CHUYỂN THỦ CÔNG
        // =========================================================================
        pnlChietTinh = new Panel
        {
            Location = new Point(25, y),
            Size = new Size(pnlContent.ClientSize.Width - 50, 140),
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
            Text = "📊  TỔNG HỢP CHIẾT TÍNH ĐỊNH MỨC VẬN CHUYỂN BỘ THEO THÔNG TƯ SỐ 38/2026/TT-BXD",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(0, 85, 170),
            AutoSize = true,
            Location = new Point(16, 8),
            BackColor = Color.Transparent
        };
        pnlChietTinh.Controls.Add(lblCard3Title);

        lblHuongDan = new Label
        {
            Text = "Vui lòng nhập Tổng cự ly vận chuyển bộ (m) ở trên hoặc bấm + Thêm đoạn vận chuyển để tính toán.",
            Location = new Point(20, 45),
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Italic),
            ForeColor = Color.FromArgb(108, 117, 125)
        };

        lblNac1 = new Label { Text = "• Nấc 1 (10m khởi điểm): ...", Location = new Point(20, 42), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };
        lblNac2 = new Label { Text = "• Nấc 2 (Mỗi 10m tiếp theo <= 300m): ...", Location = new Point(20, 68), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Regular), Visible = false };

        lblTongCong = new Label
        {
            Text = "Tổng hao phí nhân công = 0 công",
            Location = new Point(20, 100),
            AutoSize = true,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(40, 167, 69)
        };

        pnlChietTinh.Controls.Add(lblHuongDan);
        pnlChietTinh.Controls.Add(lblNac1);
        pnlChietTinh.Controls.Add(lblNac2);
        pnlChietTinh.Controls.Add(lblTongCong);
        pnlContent.Controls.Add(pnlChietTinh);
        y += 155;

        // =========================================================================
        // CARD 4: QUY ĐỔI & KẾT QUẢ
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
            Text = "Chi phí vận chuyển bộ: 0 đ / đơn vị",
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

        txtDonGiaNC.Text = _donGiaNhanCong.ToString("N0", ViVn);

        // 1. Kiểm tra cấu hình đã lưu trong kho bền vững
        var savedCfg = VanChuyenStorage.GetConfigBo(_tenVatLieu);
        string? targetMa = _maDinhMucCu ?? savedCfg?.MaDinhMuc;
        if (savedCfg != null)
        {
            if (savedCfg.SoTang > 1)
            {
                chkCaoTang.Checked = true;
                numTang.Value = Math.Max(1, Math.Min(50, savedCfg.SoTang));
            }
        }

        // 2. Nhận diện định mức
        DinhMucVCBoItem? targetDm = null;
        if (!string.IsNullOrEmpty(targetMa))
        {
            targetDm = DinhMucVanChuyenDatabase.DanhSachBo.FirstOrDefault(x => x.MaHieu == targetMa);
        }
        if (targetDm == null)
        {
            targetDm = DinhMucVanChuyenDatabase.NhanDienBo(_tenVatLieu) ?? DinhMucVanChuyenDatabase.DanhSachBo.First();
        }

        int idxDm = -1;
        for (int i = 0; i < cbDinhMuc.Items.Count; i++)
        {
            if (cbDinhMuc.Items[i] is DinhMucVCBoItem item && item.MaHieu == targetDm.MaHieu)
            {
                idxDm = i;
                break;
            }
        }
        cbDinhMuc.SelectedIndex = idxDm >= 0 ? idxDm : 0;

        // 3. Nạp danh sách các đoạn vận chuyển bộ
        dgvDoanBo.Rows.Clear();
        if (savedCfg != null && savedCfg.TongCuLyMet > 0 && savedCfg.DoanBos != null && savedCfg.DoanBos.Count > 0)
        {
            TongCuLyMet = savedCfg.TongCuLyMet;
            txtTongCuLy.Text = TongCuLyMet.ToString("0.##", ViVn);
            foreach (var db in savedCfg.DoanBos)
            {
                ThemDoanBoMoi(db.DiemDau, db.DiemCuoi, db.CuLyMet, db.HeSoDiaHinh, db.SoTang, db.NacCuLy, insertIndex: -1);
            }
        }
        else
        {
            // Mở lần đầu: DỮ LIỆU HOÀN TOÀN TRỐNG
            txtTongCuLy.Text = "";
            TongCuLyMet = 0;
            dgvDoanBo.Rows.Clear();
        }

        _suppressRecalc = false;
        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    private void OnDinhMucChanged()
    {
        if (cbDinhMuc.SelectedItem is not DinhMucVCBoItem dm) return;
        SelectedDinhMuc = dm;

        lblDvtDinhMuc.Text = $"ĐVT định mức: {dm.DonViDinhMuc}";
        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    private void OnCaoTangSettingsChanged()
    {
        bool coCaoTang = chkCaoTang.Checked;
        lblTang.Visible = coCaoTang;
        numTang.Visible = coCaoTang;
        chkThangMay.Visible = coCaoTang;

        decimal kTang = TinhHeSoTangCao((int)numTang.Value, chkThangMay.Checked);

        if (coCaoTang)
        {
            if (chkThangMay.Checked)
            {
                lblHeSoTangInfo.Text = $"* Dùng máy nâng: k_tầng = 1,00";
            }
            else
            {
                int t = (int)numTang.Value;
                if (t >= 2)
                {
                    lblHeSoTangInfo.Text = $"* Vận chuyển thủ công lên Tầng {t}: k_tầng = 1,1^{t - 1} = {kTang.ToString("0.###", ViVn)} (theo TT 38/2026/TT-BXD)";
                }
                else
                {
                    lblHeSoTangInfo.Text = "* Vận chuyển tại mặt bằng Tầng 1: k_tầng = 1,00";
                }
            }
        }
        else
        {
            lblHeSoTangInfo.Text = "";
        }

        _isUpdatingGridFromCode = true;
        foreach (DataGridViewRow r in dgvDoanBo.Rows)
        {
            int t = coCaoTang ? (int)numTang.Value : 1;
            r.Cells["SoTang"].Value = t.ToString();

            var diaHinhStr = r.Cells["DiaHinh"].Value?.ToString() ?? "";
            decimal kDiaHinh = ParseHeSoDiaHinh(diaHinhStr);
            decimal kTong = kDiaHinh * kTang;
            r.Cells["HeSoK"].Value = kTong.ToString("0.##", ViVn);
        }
        _isUpdatingGridFromCode = false;

        TinhToan();
    }

    private static decimal TinhHeSoTangCao(int tang, bool dungThangMay)
    {
        if (dungThangMay || tang <= 1) return 1.0m;
        double pow = Math.Pow(1.1, tang - 1);
        return (decimal)Math.Round(pow, 4, MidpointRounding.AwayFromZero);
    }

    private void OnTongCuLyUserSubmitted()
    {
        decimal ly = TinhCuocVCOToForm.ParseVnFlex(txtTongCuLy.Text);
        if (ly <= 0)
        {
            TongCuLyMet = 0;
            txtTongCuLy.Text = "";
            dgvDoanBo.Rows.Clear();
            CapNhatThongTinDinhMucTheoTongCuLy();
            TinhToan();
            return;
        }

        if (ly > 300m)
        {
            MessageBox.Show("Theo quy định AM.21000, định mức vận chuyển thủ công chỉ áp dụng tối đa trong phạm vi <= 300m. Phần mềm sẽ giới hạn cự ly là 300m.", "Lưu ý định mức", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ly = 300m;
        }

        TongCuLyMet = ly;
        txtTongCuLy.Text = TongCuLyMet.ToString("0.##", ViVn);

        _isUpdatingGridFromCode = true;
        dgvDoanBo.Rows.Clear();

        int tangHienTai = chkCaoTang.Checked ? (int)numTang.Value : 1;
        decimal l1 = Math.Min(TongCuLyMet, 10m);
        ThemDoanBoMoi("Bãi tập kết vật liệu", "Vị trí tập kết trung gian", l1, 1.0m, tangHienTai, "10m khởi điểm", insertIndex: -1);

        if (TongCuLyMet > 10m)
        {
            decimal l2 = TongCuLyMet - 10m;
            ThemDoanBoMoi("Vị trí tập kết trung gian", "Hiện trường thi công", l2, 1.0m, tangHienTai, "Mỗi 10m tiếp theo", insertIndex: -1);
        }
        _isUpdatingGridFromCode = false;

        CapNhatThongTinDinhMucTheoTongCuLy();
        TinhToan();
    }

    private void CapNhatTongCuLyTuBang()
    {
        decimal tong = 0;
        foreach (DataGridViewRow r in dgvDoanBo.Rows)
        {
            tong += TinhCuocVCOToForm.ParseVnFlex(r.Cells["CuLy"].Value?.ToString());
        }

        TongCuLyMet = Math.Min(300m, tong);
        txtTongCuLy.Text = TongCuLyMet > 0 ? TongCuLyMet.ToString("0.##", ViVn) : "";
        CapNhatThongTinDinhMucTheoTongCuLy();
    }

    private void CapNhatThongTinDinhMucTheoTongCuLy()
    {
        if (SelectedDinhMuc == null) return;
        var dm = SelectedDinhMuc;

        if (TongCuLyMet <= 0)
        {
            lblDmInfo.Text = "Nhập Tổng cự ly vận chuyển bộ (m) ở trên hoặc thêm các đoạn ở dưới để áp dụng định mức";
            return;
        }

        var parts = new List<string>();
        parts.Add($"10m đầu = {dm.Dm10m.ToString("0.###", ViVn)} công");
        if (TongCuLyMet > 10m)
        {
            parts.Add($"Mỗi 10m tiếp theo = {dm.DmTiepTheo.ToString("0.###", ViVn)} công/10m");
        }

        lblDmInfo.Text = $"Các định mức cần có ({dm.DonViDinhMuc}): " + string.Join("  |  ", parts);
    }

    private void ThemDoanBoMoi(string diemDau, string diemCuoi, decimal cuLy, decimal heSoDiaHinh, int soTang, string? nac = null, int insertIndex = -1)
    {
        string nacStr = nac ?? (dgvDoanBo.Rows.Count == 0 ? "10m khởi điểm" : "Mỗi 10m tiếp theo");

        decimal kTang = TinhHeSoTangCao(soTang, chkThangMay.Checked);
        decimal kTong = heSoDiaHinh * kTang;

        string diaHinhStr = "Bằng phẳng, dốc <= 7° (k=1,00)";
        if (heSoDiaHinh == 1.20m) diaHinhStr = "Dốc <= 10° (k=1,20)";
        else if (heSoDiaHinh == 1.35m) diaHinhStr = "Dốc <= 15° (k=1,35)";
        else if (heSoDiaHinh == 1.70m) diaHinhStr = "Dốc <= 20° (k=1,70)";
        else if (heSoDiaHinh == 2.00m) diaHinhStr = "Dốc <= 25° (k=2,00)";
        else if (heSoDiaHinh == 2.50m) diaHinhStr = "Trơn, lầy lún (k=2,50)";
        else if (heSoDiaHinh == 1.50m) diaHinhStr = "Gồ ghề, lởm chởm (k=1,50)";

        object[] rowData = new object[]
        {
            "1",
            nacStr,
            diemDau,
            diemCuoi,
            cuLy > 0 ? cuLy.ToString("0.##", ViVn) : "0",
            diaHinhStr,
            soTang.ToString(),
            kTong.ToString("0.##", ViVn),
            "0",
            ""
        };

        if (insertIndex >= 0 && insertIndex <= dgvDoanBo.Rows.Count)
        {
            dgvDoanBo.Rows.Insert(insertIndex, rowData);
        }
        else
        {
            dgvDoanBo.Rows.Add(rowData);
        }

        for (int i = 0; i < dgvDoanBo.Rows.Count; i++)
        {
            dgvDoanBo["STT", i].Value = (i + 1).ToString();
        }

        if (!_isUpdatingGridFromCode)
        {
            CapNhatTongCuLyTuBang();
            TinhToan();
        }
    }

    private void XoaDoanBoChon()
    {
        if (dgvDoanBo.CurrentRow != null)
        {
            dgvDoanBo.Rows.Remove(dgvDoanBo.CurrentRow);
            for (int i = 0; i < dgvDoanBo.Rows.Count; i++)
            {
                dgvDoanBo["STT", i].Value = (i + 1).ToString();
            }

            CapNhatTongCuLyTuBang();
            TinhToan();
        }
    }

    private void DgvDoanBo_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
        var colName = dgvDoanBo.Columns[e.ColumnIndex].Name;

        if (colName == "DiaHinh" || colName == "SoTang" || colName == "CuLy" || colName == "NacCuLy")
        {
            var diaHinhStr = dgvDoanBo["DiaHinh", e.RowIndex].Value?.ToString() ?? "";
            decimal kDiaHinh = ParseHeSoDiaHinh(diaHinhStr);

            int tang = 1;
            int.TryParse(dgvDoanBo["SoTang", e.RowIndex].Value?.ToString(), out tang);
            if (tang < 1) tang = 1;

            decimal kTang = TinhHeSoTangCao(tang, chkThangMay.Checked);
            decimal kTong = kDiaHinh * kTang;

            dgvDoanBo["HeSoK", e.RowIndex].Value = kTong.ToString("0.##", ViVn);
        }
    }

    private static decimal ParseHeSoDiaHinh(string val)
    {
        if (string.IsNullOrWhiteSpace(val)) return 1.0m;
        if (val.Contains("1,20") || val.Contains("1.20")) return 1.20m;
        if (val.Contains("1,35") || val.Contains("1.35")) return 1.35m;
        if (val.Contains("1,70") || val.Contains("1.70")) return 1.70m;
        if (val.Contains("2,00") || val.Contains("2.00")) return 2.00m;
        if (val.Contains("1,50") || val.Contains("1.50")) return 1.50m;
        if (val.Contains("2,50") || val.Contains("2.50")) return 2.50m;
        return 1.00m;
    }

    private void ReadDoanBosFromGrid()
    {
        DoanBos.Clear();
        foreach (DataGridViewRow row in dgvDoanBo.Rows)
        {
            if (row.IsNewRow) continue;
            var nac = row.Cells["NacCuLy"].Value?.ToString() ?? "Mỗi 10m tiếp theo";
            var dDau = row.Cells["DiemDau"].Value?.ToString() ?? "";
            var dCuoi = row.Cells["DiemCuoi"].Value?.ToString() ?? "";
            decimal cuLy = TinhCuocVCOToForm.ParseVnFlex(row.Cells["CuLy"].Value?.ToString());
            var diaHinhStr = row.Cells["DiaHinh"].Value?.ToString() ?? "";
            decimal kDiaHinh = ParseHeSoDiaHinh(diaHinhStr);
            int tang = 1;
            int.TryParse(row.Cells["SoTang"].Value?.ToString(), out tang);
            if (tang < 1) tang = 1;

            DoanBos.Add(new DoanVanChuyenBo
            {
                NacCuLy = nac,
                DiemDau = dDau,
                DiemCuoi = dCuoi,
                CuLyMet = cuLy,
                HeSoDiaHinh = kDiaHinh,
                SoTang = tang
            });
        }
    }

    private void TinhToan()
    {
        if (_suppressRecalc) return;
        if (cbDinhMuc.SelectedItem is not DinhMucVCBoItem dm) return;

        ReadDoanBosFromGrid();
        DonGiaNhanCong = TinhCuocVCOToForm.ParseVnFlex(txtDonGiaNC.Text);

        decimal tongCuLy = DoanBos.Sum(x => x.CuLyMet);
        TongCuLyMet = Math.Min(300m, tongCuLy);

        if (tongCuLy <= 0 || DoanBos.Count == 0)
        {
            lblHuongDan.Visible = true;
            lblNac1.Visible = false;
            lblNac2.Visible = false;
            lblTongCong.Text = "Tổng hao phí nhân công = 0 công";
            KetQuaCuocBo = 0;
            lblQuyDoi.Text = "Chưa có cự ly vận chuyển bộ";
            lblKetQua.Text = $"Chi phí vận chuyển bộ: 0 đ / {_donViVatLieu}";
            pnlChietTinh.Height = 110;
            return;
        }

        lblHuongDan.Visible = false;

        decimal haoPhiNac1 = 0;
        decimal haoPhiNac2 = 0;
        decimal l1Thuc = 0;
        decimal l2Thuc = 0;
        var descNac1 = new List<string>();
        var descNac2 = new List<string>();

        for (int i = 0; i < DoanBos.Count; i++)
        {
            var db = DoanBos[i];
            decimal kTang = TinhHeSoTangCao(db.SoTang, chkThangMay.Checked);
            decimal kTong = db.HeSoDiaHinh * kTang;

            decimal hpDoan = 0;
            if (db.NacCuLy.Contains("10m khởi điểm") || db.NacCuLy.Contains("Nấc 1"))
            {
                l1Thuc += db.CuLyMet;
                hpDoan = dm.Dm10m * kTong;
                haoPhiNac1 += hpDoan;
                descNac1.Add($"{db.CuLyMet.ToString("0.##", ViVn)}m (k={kTong.ToString("0.##", ViVn)})");
            }
            else
            {
                l2Thuc += db.CuLyMet;
                decimal soDoan10m = db.CuLyMet / 10m;
                hpDoan = (soDoan10m * dm.DmTiepTheo) * kTong;
                haoPhiNac2 += hpDoan;
                descNac2.Add($"{db.CuLyMet.ToString("0.##", ViVn)}m ({soDoan10m.ToString("0.##", ViVn)} đoạn 10m x k={kTong.ToString("0.##", ViVn)})");
            }

            if (i < dgvDoanBo.Rows.Count)
            {
                dgvDoanBo["HaoPhiCong", i].Value = hpDoan.ToString("0.####", ViVn);
            }
        }

        decimal tongCong = haoPhiNac1 + haoPhiNac2;

        int curY = 40;
        lblNac1.Visible = true;
        lblNac1.Location = new Point(20, curY);
        string chiTiet1 = descNac1.Count > 1 ? $" [gồm: {string.Join(" + ", descNac1)}]" : "";
        lblNac1.Text = $"• Nấc 1 (10m khởi điểm):  L = {l1Thuc.ToString("0.##", ViVn)} m{chiTiet1}  x  Đm ({dm.Dm10m.ToString("0.###", ViVn)}) = {haoPhiNac1.ToString("0.####", ViVn)} công";
        curY += 26;

        if (l2Thuc > 0)
        {
            lblNac2.Visible = true;
            lblNac2.Location = new Point(20, curY);
            decimal soDoanTiepTheo = l2Thuc / 10m;
            string chiTiet2 = descNac2.Count > 1 ? $" [gồm: {string.Join(" + ", descNac2)}]" : "";
            lblNac2.Text = $"• Nấc 2 (Mỗi 10m tiếp theo <= 300m):  L = {l2Thuc.ToString("0.##", ViVn)} m ({soDoanTiepTheo.ToString("0.##", ViVn)} đoạn 10m){chiTiet2}  x  Đm ({dm.DmTiepTheo.ToString("0.###", ViVn)}) = {haoPhiNac2.ToString("0.####", ViVn)} công";
            curY += 26;
        }
        else
        {
            lblNac2.Visible = false;
        }

        lblTongCong.Location = new Point(20, curY + 6);
        lblTongCong.Text = $"Tổng cự ly: {tongCuLy.ToString("0.##", ViVn)} m  ==>  Tổng hao phí nhân công: {tongCong.ToString("0.####", ViVn)} công / 1 {dm.DonViDinhMuc}";
        pnlChietTinh.Height = curY + 40;

        decimal gia1Dm = tongCong * DonGiaNhanCong;

        decimal heSoQuyDoi = DinhMucBocXepDatabase.TinhHeSoQuyDoi(_donViVatLieu, dm.DonViDinhMuc);
        KetQuaCuocBo = Math.Round(gia1Dm * heSoQuyDoi, 0, MidpointRounding.AwayFromZero);

        if (heSoQuyDoi != 1.0m)
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dm.DonViDinhMuc} = {gia1Dm.ToString("N0", ViVn)} đ   (Quy đổi về 1 {_donViVatLieu} x {heSoQuyDoi.ToString("0.##", ViVn)})";
        }
        else
        {
            lblQuyDoi.Text = $"Định mức tính cho: 1 {dm.DonViDinhMuc} = {gia1Dm.ToString("N0", ViVn)} đ   (Cùng đơn vị tính với vật tư)";
        }

        lblKetQua.Text = $"Chi phí vận chuyển bộ: {KetQuaCuocBo.ToString("N0", ViVn)} đ / {_donViVatLieu}";
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
