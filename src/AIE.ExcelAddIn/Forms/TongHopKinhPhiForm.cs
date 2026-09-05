using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Core.Services;
using AIE.Core.Services.LapDuToan;
using AIE.Core.Services.Shared;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Helpers;
using AIE.ExcelAddIn.Services;
using ExcelDna.Integration;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWb = Microsoft.Office.Interop.Excel.Workbook;
using Font = System.Drawing.Font;
using TextBox = System.Windows.Forms.TextBox;
using CheckBox = System.Windows.Forms.CheckBox;
using Label = System.Windows.Forms.Label;
using Button = System.Windows.Forms.Button;
using GroupBox = System.Windows.Forms.GroupBox;
using Panel = System.Windows.Forms.Panel;
using ComboBox = System.Windows.Forms.ComboBox;
using RadioButton = System.Windows.Forms.RadioButton;

namespace AIE.ExcelAddIn.Forms
{
    public class TongHopKinhPhiForm : Form
    {
        private DuToan _duToan;
        private BangTongHopKinhPhiModel _model;
        private XuatBangBieuService _xuatService;

        // TabControl chính
        private TabControl tabMain;
        private TabPage tabChiPhiXD;
        private TabPage tabTMDT;

        // Controls Tab 1: Chi phí Xây dựng (Bảng 3.8 TT 36)
        private ComboBox cboLoaiCongTrinhXD;
        private ComboBox cboPhanLoaiPhuXD;
        private TextBox txtQuyMoXD;
        private TextBox txtCPCXD;
        private TextBox txtTTXD;
        private TextBox txtTNCTTTXD;
        private TextBox txtGTGTXD;
        private TextBox txtNhaTamXD;
        private DataGridView dgvPreviewChiPhiXD;
        private Button btnChuyenSangTab2;

        private ChiPhiXayDungCalc _calcService;
        private DinhMucCPCRepository _cpcRepo;
        private DinhMucTTRepository _ttRepo;
        private decimal _tongT = 0;
        private decimal _tongNC = 0;

        // Controls Tab 2: Chế độ bảng tính (Yêu cầu 2)
        private RadioButton radBangTHDT;
        private RadioButton radBangTMDT;
        public bool LaCheDoTongMucDauTu => radBangTMDT != null && radBangTMDT.Checked;

        // Controls thông số công trình
        private ComboBox cboLoaiCongTrinh;
        private ComboBox cboCapCongTrinh;
        private ComboBox cboSoBuocThietKe;

        private TextBox txtChiPhiXD;
        private TextBox txtChiPhiNT;
        private TextBox txtChiPhiTB;
        private TextBox txtChiPhiBT;

        // CheckBoxes điều kiện áp dụng hệ số (TT 38 & BTC)
        private CheckBox chkThietBi50;
        private CheckBox chkDaKiemToan;
        private CheckBox chkThueThamTra;
        private CheckBox chkCdtTuQL;
        private CheckBox chkVungKhoKhan;
        private CheckBox chkTuyenNhieuTinh;
        private CheckBox chkCaiTao;
        private CheckBox chkLapLai;

        // Thuế VAT chung (Yêu cầu 7)
        private ComboBox cboVATChung;
        private Button btnApDungVATChung;

        // Grid chi phí (Yêu cầu 1, 3, 4, 6, 7)
        private DataGridView dgvChiPhi;

        // Labels tổng cộng
        private Label lblTongTruocThue;
        private Label lblTongGTGT;
        private Label lblTongSauThue;
        private Label lblTieuDeTongSauThue;

        // Buttons thao tác
        private Button btnThemThuVien;
        private Button btnThemTuyBien;
        private Button btnXoaChiPhi;
        private Button btnKhoiPhuc;
        private Button btnTraLaiDinhMuc;

        // Controls Bottom Action Bar
        private TableLayoutPanel pnlSummary;
        private Button btnXuatExcelChinh;
        private Button btnLuu;
        private Button btnDong;

        private bool _isUpdating = false;

        public TongHopKinhPhiForm(DuToan duToan, bool macDinhTongMucDauTu = false)
        {
            _duToan = duToan ?? new DuToan();
            _xuatService = new XuatBangBieuService();

            KhoiTaoDuLieu();
            InitializeComponent();

            if (macDinhTongMucDauTu)
            {
                if (radBangTMDT != null) radBangTMDT.Checked = true;
                if (tabMain != null && tabTMDT != null) tabMain.SelectedTab = tabTMDT;
            }

            NapDuLieuLenGiaoDien();

            // Yêu cầu 1: Mở mặc định 100% màn hình làm việc (Maximized)
            this.WindowState = FormWindowState.Maximized;
            this.Load += (s, e) =>
            {
                this.WindowState = FormWindowState.Maximized;
                this.Bounds = Screen.FromControl(this).WorkingArea;
            };
            this.Shown += (s, e) =>
            {
                this.WindowState = FormWindowState.Maximized;
            };
        }

        private void KhoiTaoDuLieu()
        {
            var db = new DatabaseManager();
            _cpcRepo = new DinhMucCPCRepository(db.Context.GetConnection());
            _ttRepo = new DinhMucTTRepository(db.Context.GetConnection());
            _calcService = new ChiPhiXayDungCalc();

            _tongT = 0;
            _tongNC = 0;
            if (_duToan.DanhSachHangMuc != null)
            {
                foreach (var hm in _duToan.DanhSachHangMuc)
                {
                    foreach (var ct in hm.DanhSachCongTac)
                    {
                        _tongT += ct.ThanhTien;
                        _tongNC += ct.ThanhTienNC;
                    }
                }
            }

            bool hasScanned = false;
            if (_duToan.BangKinhPhi != null && _duToan.BangKinhPhi.Items.Count > 0)
            {
                _model = _duToan.BangKinhPhi;
            }
            else
            {
                hasScanned = QuetDuLieuTuSheetTHChiPhiXD();
                if (!hasScanned)
                {
                    decimal gXD = _duToan.ChiPhiXD?.G ?? 0m;
                    if (gXD <= 0)
                    {
                        decimal tongT = 0;
                        if (_duToan.DanhSachHangMuc != null)
                        {
                            foreach (var hm in _duToan.DanhSachHangMuc)
                            {
                                foreach (var ct in hm.DanhSachCongTac)
                                {
                                    tongT += ct.ThanhTien;
                                }
                            }
                        }
                        gXD = tongT > 0 ? tongT : 10_000_000_000m;
                    }

                    string loaiCT = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "";
                    string capCT = !string.IsNullOrEmpty(_duToan.CapCongTrinh) ? _duToan.CapCongTrinh : "";

                    _model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                        loaiCT: !string.IsNullOrEmpty(loaiCT) ? loaiCT : "Dân dụng",
                        capCT: !string.IsNullOrEmpty(capCT) ? capCT : "Cấp III",
                        soBuocTK: _duToan.SoBuocThietKe > 0 ? _duToan.SoBuocThietKe : 2,
                        chiPhiXD: gXD,
                        chiPhiTB: _duToan.ChiPhiThietBi,
                        chiPhiBT: 0m,
                        chiPhiNhaTam: 0m
                    );
                    _model.LoaiCongTrinh = loaiCT;
                    _model.CapCongTrinh = capCT;
                    if (_duToan.SoBuocThietKe == 0) _model.SoBuocThietKe = 0;
                }
            }

            // Nếu chưa quét sheet TH_ChiPhiXD thì thử quét lại để cập nhật chính xác G_XD và G_NHA_TAM
            if (!hasScanned)
            {
                QuetDuLieuTuSheetTHChiPhiXD();
            }
        }

        /// <summary>
        /// Yêu cầu 5: Tự động quét và liên kết Chi phí xây dựng từ sheet TH_ChiPhiXD trong workbook hiện tại
        /// E12: Chi phí xây dựng trước thuế (G)
        /// E13: Thuế GTGT (GTGT)
        /// E14: Chi phí xây dựng sau thuế (Gxd)
        /// E15: Chi phí nhà tạm (LT)
        /// E16: Tổng chi phí xây dựng (GXD)
        /// </summary>
        private bool QuetDuLieuTuSheetTHChiPhiXD()
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                var wb = app?.ActiveWorkbook;
                if (wb == null) return false;

                Microsoft.Office.Interop.Excel.Worksheet wsTH = null;
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
                {
                    if (sheet.Name == "TH_ChiPhiXD")
                    {
                        wsTH = sheet;
                        break;
                    }
                }
                if (wsTH == null) return false;

                object valE12 = wsTH.Range["E12"].Value2;
                object valE13 = wsTH.Range["E13"].Value2;
                object valE15 = wsTH.Range["E15"].Value2;

                decimal gXD = valE12 != null ? Convert.ToDecimal(valE12) : 0m;
                decimal tienThueXD = valE13 != null ? Convert.ToDecimal(valE13) : 0m;
                decimal ltNhaTam = valE15 != null ? Convert.ToDecimal(valE15) : 0m;

                if (gXD > 0)
                {
                    decimal vatRate = 0.10m;
                    if (tienThueXD > 0)
                    {
                        vatRate = Math.Round(tienThueXD / gXD, 2);
                    }

                    decimal ntTruocThue = 0m;
                    if (ltNhaTam > 0)
                    {
                        ntTruocThue = Math.Round(ltNhaTam / (1m + vatRate), 0);
                    }

                    if (_model == null)
                    {
                        string loaiCT = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "";
                        string capCT = !string.IsNullOrEmpty(_duToan.CapCongTrinh) ? _duToan.CapCongTrinh : "";
                        int soBuoc = _duToan.SoBuocThietKe > 0 ? _duToan.SoBuocThietKe : 2;

                        _model = DinhMucTT38Engine.TaoBangKinhPhiMacDinh(
                            loaiCT: !string.IsNullOrEmpty(loaiCT) ? loaiCT : "Dân dụng",
                            capCT: !string.IsNullOrEmpty(capCT) ? capCT : "Cấp III",
                            soBuocTK: soBuoc,
                            chiPhiXD: gXD,
                            chiPhiTB: _duToan.ChiPhiThietBi,
                            chiPhiBT: 0m,
                            chiPhiNhaTam: ntTruocThue
                        );
                        _model.LoaiCongTrinh = loaiCT;
                        _model.CapCongTrinh = capCT;
                        if (_duToan.SoBuocThietKe == 0) _model.SoBuocThietKe = 0;
                    }
                    else
                    {
                        _model.ChiPhiXDTruocThue = gXD;
                        _model.ChiPhiNhaTamTruocThue = ntTruocThue;
                    }

                    var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
                    if (itemXD != null)
                    {
                        itemXD.GiaTriTruocThue = gXD;
                        itemXD.ThueSuatGTGT = vatRate;
                    }

                    var itemNT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
                    if (itemNT != null)
                    {
                        itemNT.GiaTriTruocThue = ntTruocThue;
                        itemNT.ThueSuatGTGT = vatRate;
                    }

                    return true;
                }
            }
            catch { }
            return false;
        }

        private void InitializeComponent()
        {
            this.Text = "Hệ thống Quản lý Tổng hợp Dự toán & Tổng mức đầu tư xây dựng (Thông tư 36/2026/TT-BXD, TT 38/2026 & Bộ Tài chính)";
            this.Size = new Size(1300, 850);
            this.MinimumSize = new Size(1100, 700);
            this.WindowState = FormWindowState.Maximized; // Yêu cầu 1: Mở full 100% màn hình
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = UIHelper.GetFont(10.5f); // Yêu cầu 4: Font Be Vietnam Pro chuẩn

            // =========================================================================
            // 1. TOP CONTAINER: THIẾT KẾ BỐ CỤC CÂN ĐỐI, KHÔNG CHE KHUẤT CHỮ (Yêu cầu 3)
            // =========================================================================
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12, 10, 12, 8),
                BackColor = Color.FromArgb(248, 250, 253)
            };

            // Khung chế độ bảng tính (Yêu cầu 2: Cho người dùng chọn Bảng 2.1 THDT hay Bảng 1.2 TMĐT)
            var grpCheDo = new GroupBox
            {
                Text = "  1. CHỌN CHẾ ĐỘ BẢNG TÍNH THEO THÔNG TƯ 36/2026/TT-BXD  ",
                Dock = DockStyle.Top,
                Height = 68,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 6, 10, 6)
            };

            var pnlRadioCheDo = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            radBangTHDT = new RadioButton
            {
                Text = "Bảng 2.1: Tổng hợp Dự toán công trình (Không gồm bồi thường GPMB, Dự phòng chi phí 5%)",
                AutoSize = true,
                Checked = true,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(10, 6, 30, 0),
                Cursor = Cursors.Hand
            };
            radBangTHDT.CheckedChanged += (s, e) => { if (radBangTHDT.Checked) ChuyenCheDoBangTinh(); };

            radBangTMDT = new RadioButton
            {
                Text = "Bảng 1.2: Tổng mức đầu tư xây dựng (Bao gồm bồi thường GPMB, Dự phòng chi phí 10%)",
                AutoSize = true,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(153, 51, 0),
                Margin = new Padding(10, 6, 10, 0),
                Cursor = Cursors.Hand
            };
            radBangTMDT.CheckedChanged += (s, e) => { if (radBangTMDT.Checked) ChuyenCheDoBangTinh(); };

            pnlRadioCheDo.Controls.Add(radBangTHDT);
            pnlRadioCheDo.Controls.Add(radBangTMDT);
            grpCheDo.Controls.Add(pnlRadioCheDo);

            // Khung thông số công trình & giá trị đầu vào (Bố cục 2 dòng chuẩn, cân đối, không lệch dòng)
            var grpThongSo = new GroupBox
            {
                Text = "  2. THÔNG SỐ CÔNG TRÌNH, GIÁ TRỊ ĐẦU VÀO & THUẾ SUẤT VAT TOÀN BẢNG  ",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(12, 10, 12, 12),
                Margin = new Padding(0, 6, 0, 4)
            };

            var tblInputs = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 8,
                RowCount = 2,
                Padding = new Padding(4, 6, 4, 6)
            };

            // Thiết lập tỷ lệ cột thông thoáng
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 1
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 1
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 2
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 2
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 3
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 3
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Nhãn 4
            tblInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); // Ô nhập 4

            tblInputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
            tblInputs.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));

            // Hàng 1: Loại CT, Cấp CT, Bước TK, Thuế VAT chung
            var lblLoai = new Label { Text = "Loại công trình:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboLoaiCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboLoaiCongTrinh.Items.AddRange(new object[] { "-- Chọn loại công trình --", "Dân dụng", "Công nghiệp", "Giao thông", "Nông nghiệp & PTNT", "Hạ tầng kỹ thuật" });
            cboLoaiCongTrinh.SelectedIndex = 0;
            cboLoaiCongTrinh.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblLoai, 0, 0);
            tblInputs.Controls.Add(cboLoaiCongTrinh, 1, 0);

            var lblCap = new Label { Text = "Cấp công trình:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboCapCongTrinh = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboCapCongTrinh.Items.AddRange(new object[] { "-- Chọn cấp công trình --", "Cấp đặc biệt", "Cấp I", "Cấp II", "Cấp III", "Cấp IV" });
            cboCapCongTrinh.SelectedIndex = 0;
            cboCapCongTrinh.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblCap, 2, 0);
            tblInputs.Controls.Add(cboCapCongTrinh, 3, 0);

            var lblBuoc = new Label { Text = "Bước thiết kế:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            cboSoBuocThietKe = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f), Margin = new Padding(0, 6, 12, 6) };
            cboSoBuocThietKe.Items.AddRange(new object[] { "-- Chọn bước thiết kế --", "1 bước (Báo cáo KT-KT)", "2 bước (TKBVTC)", "3 bước (TKKT & TKBVTC)" });
            cboSoBuocThietKe.SelectedIndex = 0;
            cboSoBuocThietKe.SelectedIndexChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);
            tblInputs.Controls.Add(lblBuoc, 4, 0);
            tblInputs.Controls.Add(cboSoBuocThietKe, 5, 0);

            var lblVAT = new Label { Text = "Thuế VAT chung:", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), Margin = new Padding(2, 0, 6, 0) };
            var pnlVAT = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0, 5, 0, 5) };
            cboVATChung = new ComboBox { Width = 75, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10.5f, FontStyle.Bold), Margin = new Padding(0, 2, 6, 0) };
            cboVATChung.Items.AddRange(new object[] { "10%", "8%", "5%", "0%" });
            cboVATChung.SelectedIndex = 0;
            btnApDungVATChung = new Button
            {
                Text = "Áp dụng",
                Size = new Size(92, 30),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0)
            };
            btnApDungVATChung.Click += BtnApDungVATChung_Click;
            pnlVAT.Controls.Add(cboVATChung);
            pnlVAT.Controls.Add(btnApDungVATChung);
            tblInputs.Controls.Add(lblVAT, 6, 0);
            tblInputs.Controls.Add(pnlVAT, 7, 0);

            // Hàng 2: Chi phí Xây dựng, Chi phí Nhà tạm, Thiết bị, Bồi thường (Live formatting)
            var lblXD = new Label { Text = "Chi phí XD G_XD (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiXD = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiXD.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiXD.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiXD);
            txtChiPhiXD.KeyPress += OnMoneyKeyPress;
            txtChiPhiXD.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblXD, 0, 1);
            tblInputs.Controls.Add(txtChiPhiXD, 1, 1);

            var lblNT = new Label { Text = "Chi phí nhà tạm (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiNT = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiNT.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiNT.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiNT);
            txtChiPhiNT.KeyPress += OnMoneyKeyPress;
            txtChiPhiNT.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblNT, 2, 1);
            tblInputs.Controls.Add(txtChiPhiNT, 3, 1);

            var lblTB = new Label { Text = "Chi phí TB G_TB (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiTB = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 12, 6) };
            txtChiPhiTB.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiTB.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiTB);
            txtChiPhiTB.KeyPress += OnMoneyKeyPress;
            txtChiPhiTB.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblTB, 4, 1);
            tblInputs.Controls.Add(txtChiPhiTB, 5, 1);

            var lblBT = new Label { Text = "Bồi thường G_BT (đ):", AutoSize = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = UIHelper.GetFont(10.5f, FontStyle.Regular), ForeColor = Color.Black, Margin = new Padding(2, 0, 6, 0) };
            txtChiPhiBT = new TextBox { Dock = DockStyle.Fill, Font = UIHelper.GetFont(10.5f), TextAlign = HorizontalAlignment.Right, Margin = new Padding(0, 6, 0, 6) };
            txtChiPhiBT.LostFocus += (s, e) => CapNhatGiaTriDauVao();
            txtChiPhiBT.TextChanged += (s, e) => FormatLiveCurrency(txtChiPhiBT);
            txtChiPhiBT.KeyPress += OnMoneyKeyPress;
            txtChiPhiBT.KeyDown += OnMoneyKeyDown;
            tblInputs.Controls.Add(lblBT, 6, 1);
            tblInputs.Controls.Add(txtChiPhiBT, 7, 1);

            grpThongSo.Controls.Add(tblInputs);

            // Khung điều kiện áp dụng hệ số (Yêu cầu 3: Lưới 4 cột x 2 hàng, không che chữ, không chèn ép)
            var grpDieuKien = new GroupBox
            {
                Text = "  3. CÁC ĐIỀU KIỆN ĐẶC THÙ ÁP DỤNG HỆ SỐ ĐIỀU CHỈNH (THÔNG TƯ 38/2026/TT-BXD & BỘ TÀI CHÍNH)  ",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 4, 10, 6),
                Margin = new Padding(0, 4, 0, 0)
            };

            var tblCheckBoxes = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 4,
                RowCount = 2,
                Padding = new Padding(2)
            };
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 28));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 21));
            tblCheckBoxes.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));

            tblCheckBoxes.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));
            tblCheckBoxes.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));

            chkThietBi50 = new CheckBox { Text = "Thiết bị ≥ 50% (k=0,7 KT, QT)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkThietBi50.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkDaKiemToan = new CheckBox { Text = "Đã kiểm toán độc lập/KTNN (k=0,5 QT)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkDaKiemToan.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkThueThamTra = new CheckBox { Text = "Yêu cầu thuê thẩm tra (k=0,5 TĐ)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkThueThamTra.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkCdtTuQL = new CheckBox { Text = "Chủ đầu tư tự QLDA (k=0,8)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkCdtTuQL.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkVungKhoKhan = new CheckBox { Text = "Vùng sâu xa, hải đảo (k=1,35)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkVungKhoKhan.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkTuyenNhieuTinh = new CheckBox { Text = "Tuyến qua nhiều tỉnh (k=1,1)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkTuyenNhieuTinh.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkCaiTao = new CheckBox { Text = "Cải tạo, sửa chữa (k=1,15 TK)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkCaiTao.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            chkLapLai = new CheckBox { Text = "Thiết kế lặp lại (k=0,36)", Dock = DockStyle.Fill, AutoSize = true, Font = UIHelper.GetFont(10f), Margin = new Padding(2) };
            chkLapLai.CheckedChanged += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            tblCheckBoxes.Controls.Add(chkThietBi50, 0, 0);
            tblCheckBoxes.Controls.Add(chkDaKiemToan, 1, 0);
            tblCheckBoxes.Controls.Add(chkThueThamTra, 2, 0);
            tblCheckBoxes.Controls.Add(chkCdtTuQL, 3, 0);

            tblCheckBoxes.Controls.Add(chkVungKhoKhan, 0, 1);
            tblCheckBoxes.Controls.Add(chkTuyenNhieuTinh, 1, 1);
            tblCheckBoxes.Controls.Add(chkCaiTao, 2, 1);
            tblCheckBoxes.Controls.Add(chkLapLai, 3, 1);

            grpDieuKien.Controls.Add(tblCheckBoxes);

            topPanel.Controls.Add(grpDieuKien);
            topPanel.Controls.Add(grpThongSo);
            topPanel.Controls.Add(grpCheDo);

            // =========================================================================
            // 2. TOOLBAR: NÚT THAO TÁC ĐỒNG BỘ CHIỀU CAO 36px, TỰ CO GIÃN ĐỘ RỘNG
            // =========================================================================
            var toolPanel = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(12, 6, 12, 6), BackColor = Color.FromArgb(238, 242, 248) };
            var pnlToolButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };

            btnThemThuVien = new Button
            {
                Text = "➕ Thêm từ Thư viện...",
                AutoSize = true,
                MinimumSize = new Size(190, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnThemThuVien.Click += BtnThemThuVien_Click;

            btnThemTuyBien = new Button
            {
                Text = "➕ Thêm dòng mới",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnThemTuyBien.Click += BtnThemTuyBien_Click;

            btnXoaChiPhi = new Button
            {
                Text = "❌ Xóa dòng chọn",
                AutoSize = true,
                MinimumSize = new Size(150, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                ForeColor = Color.DarkRed,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnXoaChiPhi.Click += BtnXoaChiPhi_Click;

            btnKhoiPhuc = new Button
            {
                Text = "↶ Khôi phục chuẩn",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnKhoiPhuc.Click += (s, e) =>
            {
                if (MessageBox.Show("Khôi phục danh mục chi phí về cấu hình chuẩn ban đầu?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _model.Items = DinhMucTT38Engine.KhoiTaoDanhSachKhoanMucChuan();
                    DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                    ChuyenCheDoBangTinh();
                }
            };

            btnTraLaiDinhMuc = new Button
            {
                Text = "🔄 Tra lại định mức",
                AutoSize = true,
                MinimumSize = new Size(160, 36),
                Height = 36,
                Padding = new Padding(12, 0, 12, 0),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnTraLaiDinhMuc.Click += (s, e) => CapNhatKhiDoiThongSo(traLaiDinhMuc: true);

            pnlToolButtons.Controls.Add(btnThemThuVien);
            pnlToolButtons.Controls.Add(btnThemTuyBien);
            pnlToolButtons.Controls.Add(btnXoaChiPhi);
            pnlToolButtons.Controls.Add(btnKhoiPhuc);
            pnlToolButtons.Controls.Add(btnTraLaiDinhMuc);
            toolPanel.Controls.Add(pnlToolButtons);

            // =========================================================================
            // 3. BOTTOM PANEL: TỔNG KẾT & CÁC NÚT XUẤT EXCEL ĐỒNG BỘ 42px (Yêu cầu 2, 3)
            // =========================================================================
            // =========================================================================
            // 3. KHỐI TỔNG KẾT (ĐƯA VÀO TAB 2 - DOCK BOTTOM DƯỚI LƯỚI CHI PHÍ)
            // =========================================================================
            pnlSummary = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.FromArgb(240, 245, 252),
                Padding = new Padding(12, 4, 12, 4)
            };
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pnlSummary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));

            lblTongTruocThue = new Label { Text = "Trước thuế: 0 đ", Dock = DockStyle.Fill, Font = UIHelper.GetFont(11f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), TextAlign = ContentAlignment.MiddleLeft };
            lblTongGTGT = new Label { Text = "Thuế GTGT: 0 đ", Dock = DockStyle.Fill, Font = UIHelper.GetFont(11f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 51, 102), TextAlign = ContentAlignment.MiddleLeft };

            var pnlGrandTotal = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            lblTongSauThue = new Label { Text = "0 đ", AutoSize = true, Font = UIHelper.GetFont(13.5f, FontStyle.Bold), ForeColor = Color.DarkRed, Margin = new Padding(0, 2, 0, 0) };
            lblTieuDeTongSauThue = new Label { Text = "TỔNG DỰ TOÁN CT:", AutoSize = true, Font = UIHelper.GetFont(12.5f, FontStyle.Bold), ForeColor = Color.DarkRed, Margin = new Padding(0, 4, 8, 0) };
            pnlGrandTotal.Controls.Add(lblTongSauThue);
            pnlGrandTotal.Controls.Add(lblTieuDeTongSauThue);

            pnlSummary.Controls.Add(lblTongTruocThue, 0, 0);
            pnlSummary.Controls.Add(lblTongGTGT, 1, 0);
            pnlSummary.Controls.Add(pnlGrandTotal, 2, 0);

            // =========================================================================
            // 4. BOTTOM PANEL: THANH TÁC VỤ CHUNG Ở ĐÁY MODAL (GỌN GÀNG, KHÔNG CHECKBOX)
            // =========================================================================
            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 54,
                Padding = new Padding(15, 8, 15, 8),
                BackColor = Color.FromArgb(245, 248, 252)
            };

            var lblBottomTip = new Label
            {
                Text = "💡 Bấm 'Xuất Excel theo lựa chọn' để chọn các sheet cần xuất sang Workbook hiện hành.",
                AutoSize = true,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UIHelper.GetFont(9.5f, FontStyle.Italic),
                ForeColor = Color.FromArgb(100, 116, 139)
            };

            var pnlActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false
            };

            btnDong = new Button
            {
                Text = "Đóng",
                Size = new Size(95, 38),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnDong.Click += (s, e) => this.Close();

            btnLuu = new Button
            {
                Text = "💾 Lưu cấu hình",
                Size = new Size(130, 38),
                BackColor = Color.FromArgb(43, 87, 154),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                Margin = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnLuu.Click += BtnLuu_Click;

            btnXuatExcelChinh = new Button
            {
                Text = "📥 Xuất Excel theo lựa chọn",
                Size = new Size(230, 38),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Margin = new Padding(0),
                Cursor = Cursors.Hand
            };
            btnXuatExcelChinh.Click += (s, e) => XuatExcelTongHop();

            pnlActions.Controls.Add(btnDong);
            pnlActions.Controls.Add(btnLuu);
            pnlActions.Controls.Add(btnXuatExcelChinh);

            bottomPanel.Controls.Add(pnlActions);
            bottomPanel.Controls.Add(lblBottomTip);

            // =========================================================================
            // 4. CENTER: DATAGRIDVIEW CO GIÃN TỰ ĐỘNG 100% (Yêu cầu 1, 3, 4, 6, 7)
            // =========================================================================
            dgvChiPhi = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, // Tự động giãn cột 100%
                Font = UIHelper.GetFont(10.5f)
            };
            UIHelper.ApplyStyle(dgvChiPhi);
            dgvChiPhi.RowTemplate.Height = 36; // Rộng rãi, chữ to rõ ràng
            dgvChiPhi.CellValueChanged += DgvChiPhi_CellValueChanged;
            dgvChiPhi.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvChiPhi.IsCurrentCellDirty)
                    dgvChiPhi.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            dgvChiPhi.DataError += (s, e) => { e.Cancel = true; }; // Chống crash ComboBox

            var colActive = new DataGridViewCheckBoxColumn { Name = "colActive", HeaderText = "Dùng", Width = 55, FillWeight = 4 };
            var colSTT = new DataGridViewTextBoxColumn { Name = "colSTT", HeaderText = "STT", Width = 65, FillWeight = 5, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };
            var colTen = new DataGridViewTextBoxColumn { Name = "colTen", HeaderText = "Nội dung khoản mục chi phí", Width = 380, FillWeight = 36 };

            var colCachTinh = new DataGridViewComboBoxColumn
            {
                Name = "colCachTinh",
                HeaderText = "Cách tính",
                Width = 125,
                FillWeight = 11,
                DataSource = new string[] { "Theo tỷ lệ %", "Tự nhập tiền" }
            };

            var colCoSo = new DataGridViewComboBoxColumn
            {
                Name = "colCoSo",
                HeaderText = "Cơ sở tính",
                Width = 130,
                FillWeight = 11,
                DataSource = new string[] { "G_XD", "G_TB", "G_XD + G_TB", "Tổng trước DP", "Toàn bộ TMĐT" }
            };

            var colTyLe = new DataGridViewTextBoxColumn { Name = "colTyLe", HeaderText = "Tỷ lệ %", Width = 80, FillWeight = 7, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
            var colHeSo = new DataGridViewTextBoxColumn { Name = "colHeSo", HeaderText = "Hệ số k", Width = 70, FillWeight = 6, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };
            var colTruocThue = new DataGridViewTextBoxColumn { Name = "colTruocThue", HeaderText = "Trước thuế (đ)", Width = 145, FillWeight = 14, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight } };

            // Yêu cầu 7: Cột VAT chọn hoặc sửa từng dòng
            var colVAT = new DataGridViewComboBoxColumn
            {
                Name = "colVAT",
                HeaderText = "VAT",
                Width = 70,
                FillWeight = 6,
                DataSource = new string[] { "10%", "8%", "5%", "0%" },
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };

            var colSauThue = new DataGridViewTextBoxColumn { Name = "colSauThue", HeaderText = "Sau thuế (đ)", Width = 150, FillWeight = 14, ReadOnly = true, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Font = UIHelper.GetFont(10.5f, FontStyle.Bold) } };
            var colKyHieu = new DataGridViewTextBoxColumn { Name = "colKyHieu", HeaderText = "Ký hiệu", Width = 70, FillWeight = 6, DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter } };

            dgvChiPhi.Columns.AddRange(colActive, colSTT, colTen, colCachTinh, colCoSo, colTyLe, colHeSo, colTruocThue, colVAT, colSauThue, colKyHieu);

            // =========================================================================
            // 5. THIẾT LẬP TAB 1: CHI PHÍ XÂY DỰNG (BẢNG 3.8 TT 36)
            // =========================================================================
            tabChiPhiXD = new TabPage
            {
                Text = "  📁 1. Chi phí Xây dựng (Bảng 3.8 TT 36)  ",
                Padding = new Padding(10),
                BackColor = Color.FromArgb(248, 250, 253),
                Font = UIHelper.GetFont(10.5f)
            };

            var pnlLeftXD = new Panel
            {
                Dock = DockStyle.Left,
                Width = 475,
                Padding = new Padding(8, 8, 12, 8),
                AutoScroll = true
            };

            var grpThongSoXD = new GroupBox
            {
                Text = "  1. Thông số phân loại công trình  ",
                Dock = DockStyle.Top,
                Height = 150,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 12, 10, 10)
            };

            var lblLoaiCTXD = new Label { Text = "Loại công trình:", Location = new Point(14, 30), Width = 150, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboLoaiCongTrinhXD = new ComboBox { Location = new Point(170, 26), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboLoaiCongTrinhXD.SelectedIndexChanged += CboLoaiCongTrinhXD_SelectedIndexChanged;

            var lblPhanLoaiXD = new Label { Text = "Phân loại chi tiết:", Location = new Point(14, 68), Width = 150, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            cboPhanLoaiPhuXD = new ComboBox { Location = new Point(170, 64), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = UIHelper.GetFont(10f) };
            cboPhanLoaiPhuXD.SelectedIndexChanged += (s, e) => TuDongTraTiLeXD();

            var lblQuyMoXD = new Label { Text = "CP XD trong TMĐT:", Location = new Point(14, 106), Width = 150, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtQuyMoXD = new TextBox { Location = new Point(170, 102), Width = 110, Font = UIHelper.GetFont(10f), Text = "15", TextAlign = HorizontalAlignment.Right };
            var lblDonViQuyMo = new Label { Text = "(tỷ đồng)", Location = new Point(288, 106), Width = 80, Font = UIHelper.GetFont(10f), ForeColor = Color.FromArgb(74, 85, 104) };
            txtQuyMoXD.TextChanged += (s, e) => TuDongTraTiLeXD();

            grpThongSoXD.Controls.AddRange(new Control[] { lblLoaiCTXD, cboLoaiCongTrinhXD, lblPhanLoaiXD, cboPhanLoaiPhuXD, lblQuyMoXD, txtQuyMoXD, lblDonViQuyMo });

            var grpTyLeXD = new GroupBox
            {
                Text = "  2. Tỷ lệ % định mức áp dụng (TT 36/2026)  ",
                Dock = DockStyle.Top,
                Height = 245,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(10, 12, 10, 10),
                Margin = new Padding(0, 10, 0, 10)
            };

            int gy = 28;
            var lblCPC = new Label { Text = "Chi phí chung (CPC):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtCPCXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), TextAlign = HorizontalAlignment.Right };
            var lblDonViCPC = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblTT = new Label { Text = "Chi phí ko XĐ KL (TT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtTTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), TextAlign = HorizontalAlignment.Right };
            var lblDonViTT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblTL = new Label { Text = "Lợi nhuận (TNCTTT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtTNCTTTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "5,5", TextAlign = HorizontalAlignment.Right };
            var lblDonViTL = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblGTGT = new Label { Text = "Thuế GTGT:", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtGTGTXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "8", TextAlign = HorizontalAlignment.Right };
            var lblDonViGTGT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            gy += 38;
            var lblNhaTamXDLabel = new Label { Text = "Chi phí nhà tạm (LT):", Location = new Point(14, gy), Width = 180, Font = UIHelper.GetFont(10f), ForeColor = Color.Black };
            txtNhaTamXD = new TextBox { Location = new Point(200, gy - 2), Width = 85, Font = UIHelper.GetFont(10f), Text = "1,1", TextAlign = HorizontalAlignment.Right };
            var lblDonViNT = new Label { Text = "%", Location = new Point(295, gy), Width = 35, Font = UIHelper.GetFont(10f) };

            grpTyLeXD.Controls.AddRange(new Control[] {
                lblCPC, txtCPCXD, lblDonViCPC,
                lblTT, txtTTXD, lblDonViTT,
                lblTL, txtTNCTTTXD, lblDonViTL,
                lblGTGT, txtGTGTXD, lblDonViGTGT,
                lblNhaTamXDLabel, txtNhaTamXD, lblDonViNT
            });

            txtCPCXD.TextChanged += (s, e) => TinhToanChiPhiXD();
            txtTTXD.TextChanged += (s, e) => TinhToanChiPhiXD();
            txtTNCTTTXD.TextChanged += (s, e) => TinhToanChiPhiXD();
            txtGTGTXD.TextChanged += (s, e) => TinhToanChiPhiXD();
            txtNhaTamXD.TextChanged += (s, e) => TinhToanChiPhiXD();

            btnChuyenSangTab2 = new Button
            {
                Text = "Tiếp tục sang Tab 2: Lập TMĐT / TH Dự toán  ➜",
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = Color.FromArgb(0, 102, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 14, 0, 0)
            };
            btnChuyenSangTab2.Click += (s, e) => { tabMain.SelectedTab = tabTMDT; };

            pnlLeftXD.Controls.Add(btnChuyenSangTab2);
            pnlLeftXD.Controls.Add(grpTyLeXD);
            pnlLeftXD.Controls.Add(grpThongSoXD);

            var pnlRightXD = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 8, 8, 8)
            };

            var lblTieuDeGridXD = new Label
            {
                Text = "BẢNG TỔNG HỢP CHI PHÍ XÂY DỰNG (BẢNG 3.8 TT 36/2026/TT-BXD)",
                Dock = DockStyle.Top,
                Height = 36,
                Font = UIHelper.GetFont(12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                TextAlign = ContentAlignment.MiddleCenter
            };

            dgvPreviewChiPhiXD = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                Font = UIHelper.GetFont(10.5f)
            };
            UIHelper.ApplyStyle(dgvPreviewChiPhiXD);
            dgvPreviewChiPhiXD.RowTemplate.Height = 36;
            dgvPreviewChiPhiXD.DefaultCellStyle.Padding = new Padding(4, 3, 4, 3);

            dgvPreviewChiPhiXD.Columns.Add("DienGiai", "Nội dung chi phí");
            dgvPreviewChiPhiXD.Columns.Add("KyHieu", "Ký hiệu");
            dgvPreviewChiPhiXD.Columns.Add("CachTinh", "Cách tính");
            dgvPreviewChiPhiXD.Columns.Add("GiaTri", "Giá trị (đồng)");
            dgvPreviewChiPhiXD.Columns[0].FillWeight = 42;
            dgvPreviewChiPhiXD.Columns[1].FillWeight = 14;
            dgvPreviewChiPhiXD.Columns[2].FillWeight = 22;
            dgvPreviewChiPhiXD.Columns[3].FillWeight = 22;
            dgvPreviewChiPhiXD.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPreviewChiPhiXD.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPreviewChiPhiXD.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvPreviewChiPhiXD.Columns[3].DefaultCellStyle.Format = "N0";

            pnlRightXD.Controls.Add(dgvPreviewChiPhiXD);
            pnlRightXD.Controls.Add(lblTieuDeGridXD);

            tabChiPhiXD.Controls.Add(pnlRightXD);
            tabChiPhiXD.Controls.Add(pnlLeftXD);

            // =========================================================================
            // 6. THIẾT LẬP TAB 2: TỔNG MỨC ĐẦU TƯ & TH DỰ TOÁN (BẢNG 1.2 & 2.1)
            // =========================================================================
            tabTMDT = new TabPage
            {
                Text = "  📊 2. Tổng mức đầu tư & TH Dự toán (Bảng 1.2 & 2.1)  ",
                Padding = new Padding(0),
                BackColor = Color.FromArgb(248, 250, 253),
                Font = UIHelper.GetFont(10.5f)
            };

            tabTMDT.Controls.Add(dgvChiPhi);
            tabTMDT.Controls.Add(pnlSummary);
            tabTMDT.Controls.Add(toolPanel);
            tabTMDT.Controls.Add(topPanel);

            // =========================================================================
            // 7. GỘP CÁC TAB VÀO TABCONTROL CHÍNH & BOTTOM BAR
            // =========================================================================
            tabMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = UIHelper.GetFont(11f, FontStyle.Bold),
                ItemSize = new Size(340, 38),
                SizeMode = TabSizeMode.Fixed,
                DrawMode = TabDrawMode.OwnerDrawFixed
            };
            tabMain.DrawItem += TabMain_DrawItem;
            tabMain.SelectedIndexChanged += (s, e) => tabMain.Invalidate();
            tabMain.TabPages.Add(tabChiPhiXD);
            tabMain.TabPages.Add(tabTMDT);

            this.Controls.Add(tabMain);
            this.Controls.Add(bottomPanel);
        }

        private void NapDuLieuLenGiaoDien()
        {
            _isUpdating = true;
            try
            {
                // Nạp dữ liệu Tab 1: Chi phí Xây dựng
                NapDuLieuTabChiPhiXD();

                // Yêu cầu 2: Để trống các ô thông số đầu vào nếu chưa chọn
                if (!string.IsNullOrEmpty(_model.LoaiCongTrinh))
                {
                    cboLoaiCongTrinh.SelectedItem = _model.LoaiCongTrinh;
                }
                if (cboLoaiCongTrinh.SelectedIndex < 0) cboLoaiCongTrinh.SelectedIndex = 0;

                if (!string.IsNullOrEmpty(_model.CapCongTrinh))
                {
                    cboCapCongTrinh.SelectedItem = _model.CapCongTrinh;
                }
                if (cboCapCongTrinh.SelectedIndex < 0) cboCapCongTrinh.SelectedIndex = 0;

                if (_model.SoBuocThietKe > 0 && _model.SoBuocThietKe <= 3)
                {
                    cboSoBuocThietKe.SelectedIndex = _model.SoBuocThietKe; // Index 0 là placeholder
                }
                else
                {
                    cboSoBuocThietKe.SelectedIndex = 0;
                }

                // Format số chuẩn Việt Nam (dấu chấm hàng nghìn)
                txtChiPhiXD.Text = UIHelper.FormatTien(_model.ChiPhiXDTruocThue);
                txtChiPhiNT.Text = UIHelper.FormatTien(_model.ChiPhiNhaTamTruocThue);
                txtChiPhiTB.Text = UIHelper.FormatTien(_model.ChiPhiTBTruocThue);
                txtChiPhiBT.Text = UIHelper.FormatTien(_model.ChiPhiBTTruocThue);

                if (_model.ChiPhiTBTruocThue > 0)
                {
                    var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                    if (itemGSTB != null) itemGSTB.IsActive = true;
                    var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                    if (itemTB != null) itemTB.IsActive = true;
                }

                chkThietBi50.Checked = _model.ThietBiTren50Pct;
                chkDaKiemToan.Checked = _model.DaKiemToanDocLap;
                chkThueThamTra.Checked = _model.YeuCauThueThamTra;
                chkCdtTuQL.Checked = _model.CdtTuQuanLy;
                chkVungKhoKhan.Checked = _model.VungKhoKhan;
                chkTuyenNhieuTinh.Checked = _model.TuyenQuaNhieuTinh;
                chkCaiTao.Checked = _model.CaiTaoSuaChua;
                chkLapLai.Checked = _model.ThietKeLapLai;

                // Đồng bộ cboVATChung với thuế suất của G_XD nếu có
                var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
                if (itemXD != null)
                {
                    if (itemXD.ThueSuatGTGT == 0.08m) cboVATChung.SelectedItem = "8%";
                    else if (itemXD.ThueSuatGTGT == 0.05m) cboVATChung.SelectedItem = "5%";
                    else if (itemXD.ThueSuatGTGT == 0m) cboVATChung.SelectedItem = "0%";
                    else cboVATChung.SelectedItem = "10%";
                }

                ChuyenCheDoBangTinh();
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void NapDuLieuTabChiPhiXD()
        {
            LoadLoaiCongTrinhXD();

            if (_duToan.ChiPhiXD != null)
            {
                txtCPCXD.Text = _duToan.ChiPhiXD.TiLeCPC.ToString("0.000").Replace('.', ',');
                txtTTXD.Text = _duToan.ChiPhiXD.TiLeTT.ToString("0.000").Replace('.', ',');
                txtTNCTTTXD.Text = _duToan.ChiPhiXD.TiLeTNCTTT.ToString("0.0").Replace('.', ',');
                txtGTGTXD.Text = _duToan.ChiPhiXD.TiLeGTGT.ToString("0").Replace('.', ',');
                txtNhaTamXD.Text = _duToan.ChiPhiXD.TiLeNhaTam.ToString("0.0").Replace('.', ',');
                TinhToanChiPhiXD();
            }
            else
            {
                TuDongTraTiLeXD();
            }
        }

        private void TabMain_DrawItem(object sender, DrawItemEventArgs e)
        {
            var tabCtrl = (TabControl)sender;
            var page = tabCtrl.TabPages[e.Index];
            var rect = tabCtrl.GetTabRect(e.Index);
            bool isSelected = (tabCtrl.SelectedIndex == e.Index);

            using var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            if (isSelected)
            {
                // Tab active: Nền xanh dương thương hiệu nổi bật
                using var brushBg = new SolidBrush(Color.FromArgb(0, 102, 204));
                e.Graphics.FillRectangle(brushBg, rect);

                // Đường viền nhấn màu vàng cam ở chân tab
                using var brushIndicator = new SolidBrush(Color.FromArgb(255, 152, 0));
                e.Graphics.FillRectangle(brushIndicator, rect.X, rect.Bottom - 4, rect.Width, 4);

                // Chữ trắng in đậm
                using var fontBold = UIHelper.GetFont(11f, FontStyle.Bold);
                using var brushText = new SolidBrush(Color.White);
                e.Graphics.DrawString(page.Text, fontBold, brushText, rect, sf);
            }
            else
            {
                // Tab inactive: Nền xám nhạt, chữ xám đen
                using var brushBg = new SolidBrush(Color.FromArgb(238, 242, 246));
                e.Graphics.FillRectangle(brushBg, rect);

                using var penBorder = new Pen(Color.FromArgb(209, 213, 219));
                e.Graphics.DrawRectangle(penBorder, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

                using var fontNormal = UIHelper.GetFont(10.5f, FontStyle.Regular);
                using var brushText = new SolidBrush(Color.FromArgb(74, 85, 104));
                e.Graphics.DrawString(page.Text, fontNormal, brushText, rect, sf);
            }
        }

        private void LoadLoaiCongTrinhXD()
        {
            cboLoaiCongTrinhXD.DataSource = null;
            cboLoaiCongTrinhXD.Items.Clear();
            cboLoaiCongTrinhXD.Items.AddRange(new object[] { "Dân dụng", "Công nghiệp", "Giao thông", "Nông nghiệp & PTNT", "Hạ tầng kỹ thuật" });

            string target = !string.IsNullOrEmpty(_duToan.LoaiCongTrinh) ? _duToan.LoaiCongTrinh : "Dân dụng";
            if (target == "Nông nghiệp và môi trường") target = "Nông nghiệp & PTNT";

            if (cboLoaiCongTrinhXD.Items.Contains(target))
            {
                cboLoaiCongTrinhXD.SelectedItem = target;
            }
            else
            {
                cboLoaiCongTrinhXD.SelectedIndex = 0;
            }
        }

        private void CboLoaiCongTrinhXD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboLoaiCongTrinhXD.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinhXD.SelectedItem.ToString();

            // Đồng bộ sang cboLoaiCongTrinh của Tab 2
            if (cboLoaiCongTrinh != null && cboLoaiCongTrinh.Items.Contains(loaiCT))
            {
                cboLoaiCongTrinh.SelectedItem = loaiCT;
            }
            _duToan.LoaiCongTrinh = loaiCT;
            if (_model != null) _model.LoaiCongTrinh = loaiCT;

            var allData = _cpcRepo.GetAll().Where(x => 
                x.LoaiCongTrinh == loaiCT || 
                (loaiCT == "Nông nghiệp & PTNT" && x.LoaiCongTrinh == "Nông nghiệp và môi trường")
            ).ToList();

            var phanLoai = allData.Where(x => !string.IsNullOrEmpty(x.PhanLoaiPhu))
                                  .Select(x => x.PhanLoaiPhu)
                                  .Distinct()
                                  .ToList();

            cboPhanLoaiPhuXD.Items.Clear();
            if (phanLoai.Count == 0)
            {
                cboPhanLoaiPhuXD.Items.Add("--- Không có ---");
                cboPhanLoaiPhuXD.Enabled = false;
            }
            else
            {
                cboPhanLoaiPhuXD.Enabled = true;
                cboPhanLoaiPhuXD.Items.Add("--- Mặc định ---");
                foreach (var item in phanLoai)
                    cboPhanLoaiPhuXD.Items.Add(item);
            }
            cboPhanLoaiPhuXD.SelectedIndex = 0;

            if (loaiCT == "Công nghiệp" || loaiCT == "Giao thông")
                txtTNCTTTXD.Text = "6,0";
            else
                txtTNCTTTXD.Text = "5,5";

            TuDongTraTiLeXD();
        }

        private void TuDongTraTiLeXD()
        {
            if (cboLoaiCongTrinhXD == null || cboLoaiCongTrinhXD.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinhXD.SelectedItem.ToString();
            string phanLoai = cboPhanLoaiPhuXD.SelectedIndex > 0 ? cboPhanLoaiPhuXD.SelectedItem.ToString() : null;

            decimal.TryParse(txtQuyMoXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal quyMo);

            // Tra CPC
            var listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
            if (listCPC.Count == 0 && phanLoai != null)
                listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null);

            if (listCPC.Count > 0)
            {
                decimal cpc = InterpolationHelper.NoiSuyTiLeCPC(listCPC, quyMo);
                txtCPCXD.Text = cpc.ToString("0.000").Replace('.', ',');
            }

            // Tra TT
            var tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
            if (tt == null && phanLoai != null)
                tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, null);

            if (tt != null)
            {
                txtTTXD.Text = tt.TiLe.ToString("0.000").Replace('.', ',');
            }

            TinhToanChiPhiXD();
        }

        private void TinhToanChiPhiXD()
        {
            if (_isUpdating) return;

            decimal.TryParse(txtCPCXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal cpc);
            decimal.TryParse(txtTTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal tt);
            decimal.TryParse(txtTNCTTTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal tncttt);
            decimal.TryParse(txtGTGTXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal gtgt);
            decimal.TryParse(txtNhaTamXD.Text.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal nhatam);

            decimal tongMay = _tongT - _tongNC - (_duToan.DanhSachHangMuc?.Sum(hm => hm.DanhSachCongTac.Sum(c => c.ThanhTienVL)) ?? 0);
            decimal tongVL = _tongT - _tongNC - tongMay;

            var kq = _calcService.Tinh(tongVL, _tongNC, tongMay, cpc, tt, tncttt, gtgt, nhatam, "T");
            _duToan.ChiPhiXD = kq;

            if (dgvPreviewChiPhiXD != null)
            {
                dgvPreviewChiPhiXD.Rows.Clear();
                dgvPreviewChiPhiXD.Rows.Add("I. Chi phí trực tiếp", "T", "VL + NC + M", kq.T);
                dgvPreviewChiPhiXD.Rows.Add("- Chi phí vật liệu", "VL", "", kq.VL);
                dgvPreviewChiPhiXD.Rows.Add("- Chi phí nhân công", "NC", "", kq.NC);
                dgvPreviewChiPhiXD.Rows.Add("- Chi phí máy", "M", "", kq.M);
                dgvPreviewChiPhiXD.Rows.Add("II. Chi phí gián tiếp", "GT", "CPC + TT", kq.GT);
                dgvPreviewChiPhiXD.Rows.Add($"- Chi phí chung ({cpc}%)", "CPC", "T × tỷ lệ", kq.CPC);
                dgvPreviewChiPhiXD.Rows.Add($"- CP không xác định KL ({tt}%)", "TT", "T × tỷ lệ", kq.TT);
                dgvPreviewChiPhiXD.Rows.Add($"III. Thu nhập chịu thuế tính trước ({tncttt}%)", "TL", "(T + GT) × tỷ lệ", kq.TL);
                dgvPreviewChiPhiXD.Rows.Add("IV. Chi phí xây dựng trước thuế", "G", "T + GT + TL", kq.G);
                dgvPreviewChiPhiXD.Rows.Add($"V. Thuế GTGT ({gtgt}%)", "GTGT", "G × tỷ lệ", kq.GTGT);
                dgvPreviewChiPhiXD.Rows.Add("VI. Chi phí xây dựng sau thuế", "Gxd", "G + GTGT", kq.Gxd);
                dgvPreviewChiPhiXD.Rows.Add($"VII. Chi phí nhà tạm ({nhatam}%)", "LT", "Gxd × tỷ lệ", kq.LT);

                var row = new DataGridViewRow();
                row.CreateCells(dgvPreviewChiPhiXD, "VIII. TỔNG CỘNG CHI PHÍ XÂY DỰNG", "GXD", "Gxd + LT", kq.GXD);
                row.DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
                row.DefaultCellStyle.BackColor = Color.LightYellow;
                dgvPreviewChiPhiXD.Rows.Add(row);

                int[] boldIndices = { 0, 4, 7, 8, 9, 10, 11 };
                foreach (int idx in boldIndices)
                {
                    if (idx < dgvPreviewChiPhiXD.Rows.Count)
                    {
                        dgvPreviewChiPhiXD.Rows[idx].DefaultCellStyle.Font = new Font(dgvPreviewChiPhiXD.Font, FontStyle.Bold);
                    }
                }

                // Cố định chiều cao 36px cho tất cả các dòng, chống bị che chữ
                foreach (DataGridViewRow r in dgvPreviewChiPhiXD.Rows)
                {
                    r.Height = 36;
                }
            }

            DongBoSangTab2(kq, gtgt, nhatam);
        }

        private void DongBoSangTab2(ChiPhiXayDung kq, decimal gtgt, decimal nhatam)
        {
            if (_model == null) return;

            decimal vatRate = gtgt / 100m;
            decimal ntTruocThue = Math.Round(kq.G * nhatam / 100m, 0);

            _model.ChiPhiXDTruocThue = kq.G;
            _model.ChiPhiNhaTamTruocThue = ntTruocThue;

            if (txtChiPhiXD != null) txtChiPhiXD.Text = UIHelper.FormatTien(kq.G);
            if (txtChiPhiNT != null) txtChiPhiNT.Text = UIHelper.FormatTien(ntTruocThue);

            if (cboVATChung != null)
            {
                string vatText = $"{gtgt:0}%";
                if (cboVATChung.Items.Contains(vatText))
                {
                    cboVATChung.SelectedItem = vatText;
                }
            }

            var itemXD = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_XD");
            if (itemXD != null)
            {
                itemXD.GiaTriTruocThue = kq.G;
                itemXD.ThueSuatGTGT = vatRate;
            }

            var itemNT = _model.Items.FirstOrDefault(x => x.MaChiPhi == "G_NHA_TAM");
            if (itemNT != null)
            {
                itemNT.GiaTriTruocThue = ntTruocThue;
                itemNT.ThueSuatGTGT = vatRate;
            }

            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
            _model.TinhToanLai();

            if (dgvChiPhi != null && dgvChiPhi.Rows.Count > 0)
            {
                HienThiDuLieuLenGrid();
            }
            CapNhatThanhTongCong();
        }

        /// <summary>
        /// Xử lý chuyển đổi giữa Bảng Tổng hợp dự toán (Bảng 2.1) và Bảng Tổng mức đầu tư (Bảng 1.2) - Yêu cầu 2
        /// </summary>
        private void ChuyenCheDoBangTinh()
        {
            if (LaCheDoTongMucDauTu)
            {
                // CHẾ ĐỘ TỔNG MỨC ĐẦU TƯ (BẢNG 1.2)
                txtChiPhiBT.Enabled = true;
                var itemBT = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
                if (itemBT != null) itemBT.IsActive = true;

                // Dự phòng TMĐT chuẩn 10%
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null) itemDP.TyLePhanTram = 10.0m;

                lblTieuDeTongSauThue.Text = "TỔNG MỨC ĐẦU TƯ:";
            }
            else
            {
                // CHẾ ĐỘ TỔNG HỢP DỰ TOÁN CÔNG TRÌNH (BẢNG 2.1)
                txtChiPhiBT.Enabled = false;
                var itemBT = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
                if (itemBT != null) itemBT.IsActive = false; // THDT không có bồi thường GPMB

                // Dự phòng THDT chuẩn 5%
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null) itemDP.TyLePhanTram = 5.0m;

                lblTieuDeTongSauThue.Text = "TỔNG DỰ TOÁN CT:";
            }

            DanhLaiSoThuTu();
            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        /// <summary>
        /// Đánh lại số thứ tự chuẩn cho toàn bộ danh mục chi phí theo chế độ TMĐT (Bảng 1.2) hoặc THDT (Bảng 2.1)
        /// </summary>
        private void DanhLaiSoThuTu()
        {
            bool laTMDT = LaCheDoTongMucDauTu;

            int sttBT = 1;
            int sttXD = 1;
            int sttTB = 1;
            int sttQLDA = 1;
            int sttTV = 1;
            int sttKhac = 1;
            int sttDP = 1;

            int countBT = _model.Items.Count(x => x.Nhom == NhomChiPhi.BoiThuong_TDC);
            int countTB = _model.Items.Count(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
            int countQLDA = _model.Items.Count(x => x.Nhom == NhomChiPhi.QuanLyDuAn);

            foreach (var item in _model.Items)
            {
                switch (item.Nhom)
                {
                    case NhomChiPhi.BoiThuong_TDC:
                        item.STT = countBT > 1 ? $"1.{sttBT++}" : "1";
                        break;

                    case NhomChiPhi.ChiPhiXayDung:
                        int prefixXD = laTMDT ? 2 : 1;
                        if (item.MaChiPhi == "G_XD") item.STT = $"{prefixXD}.1";
                        else if (item.MaChiPhi == "G_NHA_TAM") item.STT = $"{prefixXD}.2";
                        else item.STT = $"{prefixXD}.{++sttXD}";
                        break;

                    case NhomChiPhi.ChiPhiThietBi:
                        int prefixTB = laTMDT ? 3 : 2;
                        item.STT = countTB > 1 ? $"{prefixTB}.{sttTB++}" : $"{prefixTB}";
                        break;

                    case NhomChiPhi.QuanLyDuAn:
                        int prefixQLDA = laTMDT ? 4 : 3;
                        item.STT = countQLDA > 1 ? $"{prefixQLDA}.{sttQLDA++}" : $"{prefixQLDA}";
                        break;

                    case NhomChiPhi.TuVanDauTuXD:
                        int prefixTV = laTMDT ? 5 : 4;
                        item.STT = $"{prefixTV}.{sttTV++}";
                        break;

                    case NhomChiPhi.ChiPhiKhac:
                        int prefixKhac = laTMDT ? 6 : 5;
                        item.STT = $"{prefixKhac}.{sttKhac++}";
                        break;

                    case NhomChiPhi.ChiPhiDuPhong:
                        int prefixDP = laTMDT ? 7 : 6;
                        item.STT = $"{prefixDP}.{sttDP++}";
                        break;
                }
            }
        }

        private void HienThiDuLieuLenGrid()
        {
            _isUpdating = true;
            dgvChiPhi.Rows.Clear();

            // Lọc các khoản mục hiển thị:
            // 1. Nếu là THDT thì ẩn nhóm Bồi thường GPMB
            // 2. Nếu không có chi phí Thiết bị (G_TB == 0), ẩn các khoản mục phụ thuộc thiết bị mà không active (như TV_GS_TB)
            var itemsToShow = _model.Items.Where(x =>
            {
                if (!LaCheDoTongMucDauTu && x.Nhom == NhomChiPhi.BoiThuong_TDC) return false;
                if (_model.ChiPhiTBTruocThue <= 0 && x.CoSoTinh == CoSoTinhChiPhi.ChiPhiThietBi && !x.IsActive) return false;
                return true;
            }).ToList();

            foreach (var item in itemsToShow)
            {
                string cachTinhStr = item.CachTinh == CachTinhChiPhi.TheoTyLeDinhMuc ? "Theo tỷ lệ %" : "Tự nhập tiền";
                string coSoStr = "G_XD + G_TB";
                switch (item.CoSoTinh)
                {
                    case CoSoTinhChiPhi.ChiPhiXayDung: coSoStr = "G_XD"; break;
                    case CoSoTinhChiPhi.ChiPhiThietBi: coSoStr = "G_TB"; break;
                    case CoSoTinhChiPhi.TongChiPhiTruocDuPhong: coSoStr = "Tổng trước DP"; break;
                    case CoSoTinhChiPhi.TongMucDauTu: coSoStr = "Toàn bộ TMĐT"; break;
                }

                string vatStr = "10%";
                if (item.ThueSuatGTGT == 0.08m) vatStr = "8%";
                else if (item.ThueSuatGTGT == 0.05m) vatStr = "5%";
                else if (item.ThueSuatGTGT == 0m) vatStr = "0%";

                int rowIdx = dgvChiPhi.Rows.Add(
                    item.IsActive,
                    item.STT,
                    item.TenChiPhi,
                    cachTinhStr,
                    coSoStr,
                    item.TyLePhanTram > 0 ? UIHelper.FormatTyLe(item.TyLePhanTram) : "",
                    item.HeSoDieuChinh.ToString("0.00", UIHelper.ViCulture),
                    UIHelper.FormatTien(item.GiaTriTruocThue),
                    vatStr,
                    UIHelper.FormatTien(item.GiaTriSauThue),
                    item.KyHieu
                );

                var row = dgvChiPhi.Rows[rowIdx];
                row.Tag = item;

                // Định dạng nổi bật các dòng tổng nhóm chính
                if (item.Nhom == NhomChiPhi.BoiThuong_TDC || item.Nhom == NhomChiPhi.ChiPhiXayDung ||
                    item.Nhom == NhomChiPhi.ChiPhiThietBi || item.Nhom == NhomChiPhi.QuanLyDuAn ||
                    item.Nhom == NhomChiPhi.ChiPhiDuPhong)
                {
                    row.DefaultCellStyle.Font = UIHelper.GetFont(10.5f, FontStyle.Bold);
                    row.DefaultCellStyle.BackColor = Color.FromArgb(240, 245, 255);
                }

                if (!item.IsActive)
                {
                    row.DefaultCellStyle.ForeColor = Color.Gray;
                }
            }

            CapNhatThanhTongCong();
            _isUpdating = false;
        }

        private void CapNhatThanhTongCong()
        {
            lblTongTruocThue.Text = $"Trước thuế: {UIHelper.FormatTien(_model.TongTruocThue)} đ";
            lblTongGTGT.Text = $"Thuế GTGT: {UIHelper.FormatTien(_model.TongTienThueGTGT)} đ";
            lblTongSauThue.Text = $"{UIHelper.FormatTien(_model.TongSauThue)} đ";
        }

        private void CapNhatKhiDoiThongSo(bool traLaiDinhMuc)
        {
            if (_isUpdating) return;

            string loaiSel = cboLoaiCongTrinh.SelectedIndex > 0 ? cboLoaiCongTrinh.SelectedItem?.ToString() : "";
            string capSel = cboCapCongTrinh.SelectedIndex > 0 ? cboCapCongTrinh.SelectedItem?.ToString() : "";
            int buocSel = cboSoBuocThietKe.SelectedIndex > 0 ? cboSoBuocThietKe.SelectedIndex : 0;

            _model.LoaiCongTrinh = loaiSel;
            _model.CapCongTrinh = capSel;
            _model.SoBuocThietKe = buocSel;

            _model.ThietBiTren50Pct = chkThietBi50.Checked;
            _model.DaKiemToanDocLap = chkDaKiemToan.Checked;
            _model.YeuCauThueThamTra = chkThueThamTra.Checked;
            _model.CdtTuQuanLy = chkCdtTuQL.Checked;
            _model.VungKhoKhan = chkVungKhoKhan.Checked;
            _model.TuyenQuaNhieuTinh = chkTuyenNhieuTinh.Checked;
            _model.CaiTaoSuaChua = chkCaiTao.Checked;
            _model.ThietKeLapLai = chkLapLai.Checked;

            if (traLaiDinhMuc && !string.IsNullOrEmpty(loaiSel) && !string.IsNullOrEmpty(capSel))
            {
                DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                var itemDP = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (itemDP != null)
                {
                    itemDP.TyLePhanTram = LaCheDoTongMucDauTu ? 10.0m : 5.0m;
                }
            }

            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        private void CapNhatGiaTriDauVao()
        {
            if (_isUpdating) return;

            _model.ChiPhiXDTruocThue = UIHelper.ParseTien(txtChiPhiXD.Text);
            _model.ChiPhiNhaTamTruocThue = UIHelper.ParseTien(txtChiPhiNT.Text);
            _model.ChiPhiTBTruocThue = UIHelper.ParseTien(txtChiPhiTB.Text);
            _model.ChiPhiBTTruocThue = UIHelper.ParseTien(txtChiPhiBT.Text);

            // Format lại đẹp chuẩn dấu chấm
            txtChiPhiXD.Text = UIHelper.FormatTien(_model.ChiPhiXDTruocThue);
            txtChiPhiNT.Text = UIHelper.FormatTien(_model.ChiPhiNhaTamTruocThue);
            txtChiPhiTB.Text = UIHelper.FormatTien(_model.ChiPhiTBTruocThue);
            txtChiPhiBT.Text = UIHelper.FormatTien(_model.ChiPhiBTTruocThue);

            // Tự động kích hoạt chi phí thiết bị và chi phí giám sát lắp đặt thiết bị (TV_GS_TB) khi chi phí thiết bị > 0
            var itemTB = _model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
            if (itemTB != null)
            {
                itemTB.GiaTriTruocThue = _model.ChiPhiTBTruocThue;
                itemTB.IsActive = _model.ChiPhiTBTruocThue > 0;
            }

            var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
            if (itemGSTB != null)
            {
                itemGSTB.IsActive = _model.ChiPhiTBTruocThue > 0;
            }

            // Cập nhật lại định mức cho các khoản mục phụ thuộc G_TB và G_XD + G_TB
            DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);

            _model.TinhToanLai();
            HienThiDuLieuLenGrid();
        }

        /// <summary>
        /// Yêu cầu 7: Áp dụng thuế VAT chung cho toàn bộ bảng
        /// </summary>
        private void BtnApDungVATChung_Click(object sender, EventArgs e)
        {
            string vatSel = cboVATChung.SelectedItem?.ToString() ?? "10%";
            decimal newVAT = 0.10m;
            if (vatSel == "8%") newVAT = 0.08m;
            else if (vatSel == "5%") newVAT = 0.05m;
            else if (vatSel == "0%") newVAT = 0m;

            if (MessageBox.Show($"Bạn có muốn áp dụng mức thuế VAT {vatSel} cho toàn bộ các khoản mục chi phí trên bảng?\n(Lưu ý: Các khoản phí ngân sách nhà nước như Phí thẩm định sẽ tự động giữ nguyên 0%)", "Xác nhận áp dụng VAT", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (var item in _model.Items)
                {
                    // Giữ 0% cho phí ngân sách nhà nước
                    if (item.MaChiPhi == "K_TD_DA" || item.MaChiPhi == "K_TD_TK" || item.MaChiPhi == "K_TD_DT" || item.MaChiPhi == "K_TT_QUYETTOAN")
                    {
                        continue;
                    }
                    item.ThueSuatGTGT = newVAT;
                }

                _model.TinhToanLai();
                HienThiDuLieuLenGrid();
            }
        }

        private void DgvChiPhi_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isUpdating || e.RowIndex < 0) return;

            var row = dgvChiPhi.Rows[e.RowIndex];
            if (row.Tag is not ChiPhiKinhPhiItem item) return;

            string colName = dgvChiPhi.Columns[e.ColumnIndex].Name;

            if (colName == "colActive")
            {
                item.IsActive = Convert.ToBoolean(row.Cells["colActive"].Value ?? true);
                row.DefaultCellStyle.ForeColor = item.IsActive ? Color.Black : Color.Gray;
                _model.TinhToanLai();
                CapNhatThanhTongCong();
                return;
            }

            if (colName == "colTen")
            {
                item.TenChiPhi = row.Cells["colTen"].Value?.ToString() ?? "";
            }
            else if (colName == "colCachTinh")
            {
                string ct = row.Cells["colCachTinh"].Value?.ToString() ?? "";
                item.CachTinh = ct.Contains("tỷ lệ") ? CachTinhChiPhi.TheoTyLeDinhMuc : CachTinhChiPhi.NhapTruocThue;
            }
            else if (colName == "colCoSo")
            {
                string cs = row.Cells["colCoSo"].Value?.ToString() ?? "";
                if (cs == "G_XD") item.CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung;
                else if (cs == "G_TB") item.CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi;
                else if (cs == "Tổng trước DP") item.CoSoTinh = CoSoTinhChiPhi.TongChiPhiTruocDuPhong;
                else if (cs == "Toàn bộ TMĐT") item.CoSoTinh = CoSoTinhChiPhi.TongMucDauTu;
                else item.CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi;
            }
            else if (colName == "colTyLe")
            {
                item.TyLePhanTram = UIHelper.ParseTyLe(row.Cells["colTyLe"].Value?.ToString() ?? "0");
            }
            else if (colName == "colHeSo")
            {
                item.HeSoDieuChinh = UIHelper.ParseTyLe(row.Cells["colHeSo"].Value?.ToString() ?? "1");
            }
            else if (colName == "colTruocThue")
            {
                decimal tt = UIHelper.ParseTien(row.Cells["colTruocThue"].Value?.ToString() ?? "0");
                item.GiaTriTruocThue = tt;
                item.CachTinh = CachTinhChiPhi.NhapTruocThue;
                if (item.MaChiPhi == "G_XD")
                {
                    _model.ChiPhiXDTruocThue = tt;
                    txtChiPhiXD.Text = UIHelper.FormatTien(tt);
                }
                else if (item.MaChiPhi == "G_NHA_TAM")
                {
                    _model.ChiPhiNhaTamTruocThue = tt;
                    txtChiPhiNT.Text = UIHelper.FormatTien(tt);
                }
                else if (item.Nhom == NhomChiPhi.ChiPhiThietBi)
                {
                    _model.ChiPhiTBTruocThue = tt;
                    txtChiPhiTB.Text = UIHelper.FormatTien(tt);
                    var itemGSTB = _model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                    if (itemGSTB != null)
                    {
                        itemGSTB.IsActive = tt > 0;
                    }
                    DinhMucTT38Engine.CapNhatToanBoDinhMucVaTinhToan(_model);
                }
                else if (item.Nhom == NhomChiPhi.BoiThuong_TDC)
                {
                    _model.ChiPhiBTTruocThue = tt;
                    txtChiPhiBT.Text = UIHelper.FormatTien(tt);
                }
            }
            else if (colName == "colVAT")
            {
                // Yêu cầu 7: Cho phép sửa thuế VAT theo từng dòng riêng lẻ
                string vatStr = row.Cells["colVAT"].Value?.ToString() ?? "10%";
                if (vatStr.Contains("8")) item.ThueSuatGTGT = 0.08m;
                else if (vatStr.Contains("5")) item.ThueSuatGTGT = 0.05m;
                else if (vatStr.Contains("0")) item.ThueSuatGTGT = 0m;
                else item.ThueSuatGTGT = 0.10m;
            }
            else if (colName == "colKyHieu")
            {
                item.KyHieu = row.Cells["colKyHieu"].Value?.ToString() ?? "";
            }

            _model.TinhToanLai();
            _isUpdating = true;
            row.Cells["colTruocThue"].Value = UIHelper.FormatTien(item.GiaTriTruocThue);
            row.Cells["colSauThue"].Value = UIHelper.FormatTien(item.GiaTriSauThue);
            _isUpdating = false;
            CapNhatThanhTongCong();
        }

        private void BtnThemThuVien_Click(object sender, EventArgs e)
        {
            using var dlg = new ChonChiPhiThuVienForm();
            if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedItems.Count > 0)
            {
                int insertIdx = _model.Items.FindIndex(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
                if (insertIdx < 0) insertIdx = _model.Items.Count;

                foreach (var item in dlg.SelectedItems)
                {
                    if (_model.Items.Any(x => x.MaChiPhi == item.MaChiPhi))
                    {
                        item.MaChiPhi += "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    }
                    _model.Items.Insert(insertIdx++, item);
                }

                DanhLaiSoThuTu();
                _model.TinhToanLai();
                HienThiDuLieuLenGrid();
            }
        }

        private void BtnThemTuyBien_Click(object sender, EventArgs e)
        {
            var newItem = new ChiPhiKinhPhiItem
            {
                MaChiPhi = "K_CUSTOM_" + (_model.Items.Count + 1),
                TenChiPhi = "Chi phí mới (Nhấp đôi để sửa tên)",
                Nhom = NhomChiPhi.ChiPhiKhac,
                CachTinh = CachTinhChiPhi.NhapTruocThue,
                CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                GiaTriTruocThue = 0m,
                ThueSuatGTGT = 0.10m,
                KyHieu = "Gk" + (_model.Items.Count(x => x.Nhom == NhomChiPhi.ChiPhiKhac) + 1),
                IsActive = true,
                IsUserAdded = true
            };

            int insertIdx = _model.Items.FindIndex(x => x.Nhom == NhomChiPhi.ChiPhiDuPhong);
            if (insertIdx < 0) insertIdx = _model.Items.Count;

            _model.Items.Insert(insertIdx, newItem);
            DanhLaiSoThuTu();
            HienThiDuLieuLenGrid();
        }

        private void BtnXoaChiPhi_Click(object sender, EventArgs e)
        {
            if (dgvChiPhi.CurrentRow?.Tag is not ChiPhiKinhPhiItem item) return;

            if (item.IsReadOnly)
            {
                MessageBox.Show($"Khoản mục [{item.TenChiPhi}] là chi phí cốt lõi của bảng tính, không được xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show($"Bạn có chắc muốn xóa khoản mục [{item.TenChiPhi}]?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _model.Items.Remove(item);
                DanhLaiSoThuTu();
                _model.TinhToanLai();
                HienThiDuLieuLenGrid();
            }
        }

        private void BtnLuu_Click(object sender, EventArgs e)
        {
            _duToan.BangKinhPhi = _model;
            _duToan.LoaiCongTrinh = _model.LoaiCongTrinh;
            _duToan.CapCongTrinh = _model.CapCongTrinh;

            MessageBox.Show("Đã lưu thiết lập Bảng tổng hợp kinh phí thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void XuatExcelTongHop()
        {
            try
            {
                var app = (ExcelApp)ExcelDnaUtil.Application;
                var wb = app.ActiveWorkbook;
                if (wb == null)
                {
                    MessageBox.Show("Không tìm thấy Workbook Excel nào đang mở.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Đảm bảo dữ liệu Chi phí XD đã được tính
                if (_duToan.ChiPhiXD == null)
                {
                    TinhToanChiPhiXD();
                }

                _duToan.BangKinhPhi = _model;

                // Mở hộp thoại chọn các bảng biểu cần xuất
                using var dialog = new ChonBangXuatExcelDialog(isCheDoTMDT: LaCheDoTongMucDauTu);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var opts = dialog.LuaChon;
                if (opts == null || !opts.CoItNhatMotBangDuocChon())
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một bảng biểu để xuất Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var loading = new LoadingForm("Đang xuất các bảng biểu ra Excel...");
                loading.Show();
                System.Windows.Forms.Application.DoEvents();

                // Thực hiện xuất các bảng đã chọn
                _xuatService.XuatCacBangTheoTuyChon(wb, _duToan, opts);

                loading.Close();
                loading.Dispose();

                var dsDaXuat = opts.GetDanhSachDaChon().Select(x => $"- {x}").ToList();
                MessageBox.Show($"Đã xuất thành công các bảng biểu sang Excel:\n{string.Join("\n", dsDaXuat)}", "Xuất Excel Thành Công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool _isFormattingText = false;

        /// <summary>
        /// Yêu cầu 6: Định dạng số thời gian thực chuẩn Việt Nam (phân cách hàng nghìn bằng dấu chấm) khi người dùng gõ phím
        /// Giữ nguyên vị trí con trỏ chuột không bị nhảy về cuối ô.
        /// </summary>
        private void FormatLiveCurrency(TextBox tb)
        {
            if (_isUpdating || _isFormattingText || string.IsNullOrWhiteSpace(tb.Text)) return;

            try
            {
                _isFormattingText = true;
                string raw = tb.Text;
                int selStart = tb.SelectionStart;
                int rightOffset = raw.Length - selStart;

                string integerPart = raw;
                string decimalPart = "";
                int commaIndex = raw.IndexOf(',');
                if (commaIndex >= 0)
                {
                    integerPart = raw.Substring(0, commaIndex);
                    decimalPart = raw.Substring(commaIndex);
                    string decDigits = new string(decimalPart.Skip(1).Where(char.IsDigit).ToArray());
                    decimalPart = "," + decDigits;
                }

                string cleanInt = new string(integerPart.Where(char.IsDigit).ToArray());
                if (string.IsNullOrEmpty(cleanInt))
                {
                    if (commaIndex >= 0)
                    {
                        tb.Text = "0" + decimalPart;
                        tb.SelectionStart = Math.Min(tb.Text.Length, 1);
                    }
                    return;
                }

                if (decimal.TryParse(cleanInt, out decimal intVal))
                {
                    string formattedInt = intVal.ToString("#,##0", UIHelper.ViCulture);
                    string formattedTotal = formattedInt + decimalPart;

                    if (tb.Text != formattedTotal)
                    {
                        tb.Text = formattedTotal;
                        int newPos = formattedTotal.Length - rightOffset;
                        if (newPos < 0) newPos = 0;
                        if (newPos > formattedTotal.Length) newPos = formattedTotal.Length;
                        tb.SelectionStart = newPos;
                    }
                }
            }
            finally
            {
                _isFormattingText = false;
            }
        }

        private void OnMoneyKeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;

            // Chuyển đổi phím chấm (numpad) thành phẩy theo chuẩn Việt Nam
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }

            if (e.KeyChar == ',')
            {
                if (sender is TextBox tb && tb.Text.Contains(","))
                {
                    e.Handled = true; // Chỉ cho phép 1 dấu phẩy
                }
                return;
            }

            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Chỉ cho phép số
            }
        }

        private void OnMoneyKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CapNhatGiaTriDauVao();
                e.SuppressKeyPress = true;
            }
        }
    }
}
