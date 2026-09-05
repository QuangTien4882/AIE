using System;
using System.Drawing;
using System.Windows.Forms;
using AIE.ExcelAddIn.Services;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms
{
    public class ChonBangXuatExcelDialog : Form
    {
        private CheckBox chkTMDT;
        private CheckBox chkTHDT;
        private CheckBox chkTHCPXD;

        private CheckBox chkDuToan;
        private CheckBox chkPhanTich;
        private CheckBox chkTHVL;
        private CheckBox chkTHNC;
        private CheckBox chkTHMay;
        private CheckBox chkCuocVC;
        private CheckBox chkHeSo;

        private Button btnChonTatCa;
        private Button btnBoChonTatCa;
        private Button btnMacDinh;

        private Button btnXuat;
        private Button btnDong;

        private readonly bool _isCheDoTMDT;

        public LuaChonXuatExcel LuaChon { get; private set; }

        public ChonBangXuatExcelDialog(bool isCheDoTMDT = true)
        {
            _isCheDoTMDT = isCheDoTMDT;
            LuaChon = new LuaChonXuatExcel();

            InitializeComponent();
            ApDungMacDinh();
        }

        private void InitializeComponent()
        {
            this.Text = "Tùy chọn Bảng biểu Xuất sang Excel (AIE Dự Toán)";
            this.Size = new Size(680, 680);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.BackColor = Color.FromArgb(248, 250, 253);
            this.Font = UIHelper.GetFont(10f);

            // =========================================================================
            // 1. HEADER BANNER
            // =========================================================================
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 65,
                BackColor = Color.FromArgb(16, 124, 65), // Excel Green
                Padding = new Padding(20, 10, 20, 10)
            };

            var lblHeaderTitle = new Label
            {
                Text = "📥 TÙY CHỌN BẢNG BIỂU XUẤT SANG EXCEL",
                Dock = DockStyle.Top,
                Height = 26,
                Font = UIHelper.GetFont(12f, FontStyle.Bold),
                ForeColor = Color.White
            };

            var lblHeaderSub = new Label
            {
                Text = "Tích chọn các bảng biểu dự toán & tổng hợp kinh phí bạn muốn tạo/cập nhật vào Workbook:",
                Dock = DockStyle.Bottom,
                Height = 20,
                Font = UIHelper.GetFont(9.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(230, 244, 234)
            };

            pnlHeader.Controls.Add(lblHeaderTitle);
            pnlHeader.Controls.Add(lblHeaderSub);
            this.Controls.Add(pnlHeader);

            // =========================================================================
            // 2. TOOLBAR NÚT CHỌN NHANH
            // =========================================================================
            var pnlToolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(15, 6, 15, 4),
                BackColor = Color.FromArgb(240, 244, 248)
            };

            btnChonTatCa = new Button
            {
                Text = "☑ Chọn tất cả",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnChonTatCa.Click += (s, e) => SetAllCheckboxes(true);

            btnBoChonTatCa = new Button
            {
                Text = "☐ Bỏ chọn tất cả",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnBoChonTatCa.Click += (s, e) => SetAllCheckboxes(false);

            btnMacDinh = new Button
            {
                Text = "↺ Mặc định đề xuất",
                AutoSize = true,
                Height = 30,
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 8, 0)
            };
            btnMacDinh.Click += (s, e) => ApDungMacDinh();

            pnlToolbar.Controls.Add(btnChonTatCa);
            pnlToolbar.Controls.Add(btnBoChonTatCa);
            pnlToolbar.Controls.Add(btnMacDinh);
            this.Controls.Add(pnlToolbar);

            // =========================================================================
            // 3. BOTTOM PANEL: NÚT XUẤT & HỦY
            // =========================================================================
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.White,
                Padding = new Padding(15, 10, 20, 10)
            };

            var pnlBottomActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true
            };

            btnDong = new Button
            {
                Text = "Đóng",
                Size = new Size(100, 38),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10f),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 10, 0)
            };
            btnDong.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnXuat = new Button
            {
                Text = "📥 Bắt đầu xuất Excel",
                Size = new Size(200, 38),
                BackColor = Color.FromArgb(16, 124, 65),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = UIHelper.GetFont(10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0)
            };
            btnXuat.Click += BtnXuat_Click;

            pnlBottomActions.Controls.Add(btnDong);
            pnlBottomActions.Controls.Add(btnXuat);
            pnlBottom.Controls.Add(pnlBottomActions);
            this.Controls.Add(pnlBottom);

            // =========================================================================
            // 4. MAIN CONTENT PANEL CHỨA 2 GROUPBOX CHECKBOXES
            // =========================================================================
            var pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(20, 10, 20, 10)
            };

            // GROUP 1: BẢNG TỔNG HỢP KINH PHÍ (TT 36/2026/TT-BXD)
            var grpTongHop = new GroupBox
            {
                Text = "  1. BẢNG TỔNG HỢP KINH PHÍ (THÔNG TƯ 36/2026/TT-BXD)  ",
                Dock = DockStyle.Top,
                Height = 150,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(15, 12, 15, 10),
                Margin = new Padding(0, 0, 0, 10)
            };

            var pnlGrp1 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = false
            };

            chkTMDT = CreateCheckbox("Bảng 1.2: Tổng mức đầu tư xây dựng", "TongMucDauTu", Color.FromArgb(153, 51, 0));
            chkTHDT = CreateCheckbox("Bảng 2.1: Tổng hợp dự toán công trình", "TH_DuToan", Color.FromArgb(0, 102, 204));
            chkTHCPXD = CreateCheckbox("Bảng 3.8: Bảng tổng hợp chi phí xây dựng", "TH_ChiPhiXD", Color.FromArgb(0, 102, 0));

            pnlGrp1.Controls.Add(chkTMDT);
            pnlGrp1.Controls.Add(chkTHDT);
            pnlGrp1.Controls.Add(chkTHCPXD);
            grpTongHop.Controls.Add(pnlGrp1);

            // GROUP 2: BẢNG BIỂU KỸ THUẬT & DỰ TOÁN CHI TIẾT
            var grpKyThuat = new GroupBox
            {
                Text = "  2. HỆ THỐNG BẢNG BIỂU KỸ THUẬT & DỰ TOÁN CHI TIẾT  ",
                Dock = DockStyle.Top,
                Height = 295,
                Font = UIHelper.GetFont(10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102),
                Padding = new Padding(15, 12, 15, 10),
                Margin = new Padding(0, 10, 0, 0)
            };

            var pnlGrp2 = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = false
            };

            chkDuToan = CreateCheckbox("Dự toán chi phí xây dựng công trình", "DuToan", Color.Black);
            chkPhanTich = CreateCheckbox("Bảng phân tích đơn giá chi tiết", "PhanTich_DonGia", Color.Black);
            chkTHVL = CreateCheckbox("Bảng tổng hợp chênh lệch giá vật liệu", "TH_VatLieu", Color.Black);
            chkTHNC = CreateCheckbox("Bảng tổng hợp và chênh lệch nhân công", "TH_NhanCong", Color.Black);
            chkTHMay = CreateCheckbox("Bảng tổng hợp và chênh lệch máy thi công", "TH_CaMay", Color.Black);
            chkCuocVC = CreateCheckbox("Bảng chiết tính cước vận chuyển", "ChietTinh_CuocVC", Color.Black);
            chkHeSo = CreateCheckbox("Bảng xác định hệ số điều chỉnh", "HeSo_DieuChinh", Color.Black);

            pnlGrp2.Controls.Add(chkDuToan);
            pnlGrp2.Controls.Add(chkPhanTich);
            pnlGrp2.Controls.Add(chkTHVL);
            pnlGrp2.Controls.Add(chkTHNC);
            pnlGrp2.Controls.Add(chkTHMay);
            pnlGrp2.Controls.Add(chkCuocVC);
            pnlGrp2.Controls.Add(chkHeSo);
            grpKyThuat.Controls.Add(pnlGrp2);

            pnlContent.Controls.Add(grpKyThuat);
            pnlContent.Controls.Add(grpTongHop);

            this.Controls.Add(pnlContent);

            // Reorder dock hierarchy
            pnlHeader.BringToFront();
            pnlToolbar.BringToFront();
            pnlBottom.SendToBack();
            pnlContent.BringToFront();
        }

        private CheckBox CreateCheckbox(string title, string sheetName, Color tagColor)
        {
            var chk = new CheckBox
            {
                Text = $"{title}  (sheet '{sheetName}')",
                AutoSize = true,
                Font = UIHelper.GetFont(10f, FontStyle.Regular),
                ForeColor = tagColor,
                Margin = new Padding(4, 5, 4, 5),
                Cursor = Cursors.Hand
            };
            return chk;
        }

        private void SetAllCheckboxes(bool isChecked)
        {
            chkTMDT.Checked = isChecked;
            chkTHDT.Checked = isChecked;
            chkTHCPXD.Checked = isChecked;
            chkDuToan.Checked = isChecked;
            chkPhanTich.Checked = isChecked;
            chkTHVL.Checked = isChecked;
            chkTHNC.Checked = isChecked;
            chkTHMay.Checked = isChecked;
            chkCuocVC.Checked = isChecked;
            chkHeSo.Checked = isChecked;
        }

        private void ApDungMacDinh()
        {
            // Mặc định chọn bảng tổng hợp theo chế độ đang mở
            if (_isCheDoTMDT)
            {
                chkTMDT.Checked = true;
                chkTHDT.Checked = false;
            }
            else
            {
                chkTMDT.Checked = false;
                chkTHDT.Checked = true;
            }

            // Bảng tổng hợp chi phí xây dựng luôn mặc định chọn
            chkTHCPXD.Checked = true;

            // Bảng kỹ thuật cơ bản mặc định chọn
            chkDuToan.Checked = true;
            chkPhanTich.Checked = true;
            chkTHVL.Checked = true;
            chkTHNC.Checked = true;
            chkTHMay.Checked = true;
            chkCuocVC.Checked = true;
            chkHeSo.Checked = true;
        }

        private void BtnXuat_Click(object sender, EventArgs e)
        {
            LuaChon.XuatTongMucDauTu = chkTMDT.Checked;
            LuaChon.XuatTongHopDuToan = chkTHDT.Checked;
            LuaChon.XuatChiPhiXayDung = chkTHCPXD.Checked;
            LuaChon.XuatDuToanChiTiet = chkDuToan.Checked;
            LuaChon.XuatPhanTichDonGia = chkPhanTich.Checked;
            LuaChon.XuatTongHopVatLieu = chkTHVL.Checked;
            LuaChon.XuatTongHopNhanCong = chkTHNC.Checked;
            LuaChon.XuatTongHopCaMay = chkTHMay.Checked;
            LuaChon.XuatChietTinhCuocVC = chkCuocVC.Checked;
            LuaChon.XuatHeSoDieuChinh = chkHeSo.Checked;

            if (!LuaChon.CoItNhatMotBangDuocChon())
            {
                MessageBox.Show("Vui lòng tích chọn ít nhất một bảng biểu để xuất sang Excel.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
