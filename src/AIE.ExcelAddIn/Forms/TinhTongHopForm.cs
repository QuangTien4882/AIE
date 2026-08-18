using System;
using System.Globalization;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AIE.Core.Models;
using AIE.Core.Services.LapDuToan;
using AIE.Core.Services.Shared;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Services;

namespace AIE.ExcelAddIn.Forms
{
    public class TinhTongHopForm : Form
    {
        private DuToan _duToan;
        private LapDuToanExcelService _excelService;
        
        private ComboBox cboLoaiCongTrinh;
        private ComboBox cboPhanLoaiPhu;
        private TextBox txtQuyMo;
        
        private TextBox txtCPC;
        private TextBox txtTT;
        private TextBox txtTNCTTT;
        private TextBox txtGTGT;
        private TextBox txtNhaTam;
        
        private DataGridView dgvPreview;
        private Button btnXuatExcel;
        
        private ChiPhiXayDungCalc _calcService;
        private DinhMucCPCRepository _cpcRepo;
        private DinhMucTTRepository _ttRepo;

        private decimal _tongT = 0;
        private decimal _tongNC = 0;

        public TinhTongHopForm(DuToan duToan, LapDuToanExcelService excelService)
        {
            _duToan = duToan;
            _excelService = excelService;
            
            var db = new DatabaseManager();
            _cpcRepo = new DinhMucCPCRepository(db.Context.GetConnection());
            _ttRepo = new DinhMucTTRepository(db.Context.GetConnection());
            _calcService = new ChiPhiXayDungCalc();

            // Tính tổng T và NC
            foreach (var hm in _duToan.DanhSachHangMuc)
            {
                foreach (var ct in hm.DanhSachCongTac)
                {
                    _tongT += ct.ThanhTien;
                    _tongNC += ct.ThanhTienNC;
                }
            }

            InitializeComponent();
            LoadLoaiCongTrinh();
        }

        private void InitializeComponent()
        {
            this.Text = "Tính Tổng hợp Dự Toán (TT 36/2026/TT-BXD)";
            this.Size = new Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f);

            var leftPanel = new Panel { Dock = DockStyle.Left, Width = 400, Padding = new Padding(10) };
            var rightPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            this.Controls.Add(rightPanel);
            this.Controls.Add(leftPanel);

            int y = 10;
            leftPanel.Controls.Add(new Label { Text = "Loại công trình:", Location = new Point(10, y), Width = 150 });
            cboLoaiCongTrinh = new ComboBox { Location = new Point(170, y), Width = 210, DropDownStyle = ComboBoxStyle.DropDownList };
            cboLoaiCongTrinh.SelectedIndexChanged += CboLoaiCongTrinh_SelectedIndexChanged;
            leftPanel.Controls.Add(cboLoaiCongTrinh);

            y += 35;
            leftPanel.Controls.Add(new Label { Text = "Phân loại chi tiết:", Location = new Point(10, y), Width = 150 });
            cboPhanLoaiPhu = new ComboBox { Location = new Point(170, y), Width = 210, DropDownStyle = ComboBoxStyle.DropDownList };
            cboPhanLoaiPhu.SelectedIndexChanged += (s, e) => TuDongTraTiLe();
            leftPanel.Controls.Add(cboPhanLoaiPhu);

            y += 35;
            leftPanel.Controls.Add(new Label { Text = "CP XD trong TMĐT (tỷ):", Location = new Point(10, y), Width = 150 });
            txtQuyMo = new TextBox { Location = new Point(170, y), Width = 100, Text = "15" };
            txtQuyMo.TextChanged += (s, e) => TuDongTraTiLe();
            leftPanel.Controls.Add(txtQuyMo);
            
            y += 40;
            var grp = new GroupBox { Text = "Tỷ lệ % áp dụng", Location = new Point(10, y), Size = new Size(320, 220) };
            leftPanel.Controls.Add(grp);

            int gy = 25;
            grp.Controls.Add(new Label { Text = "Chi phí chung (CPC):", Location = new Point(10, gy), Width = 160 });
            txtCPC = new TextBox { Location = new Point(180, gy), Width = 60 };
            grp.Controls.Add(txtCPC);
            grp.Controls.Add(new Label { Text = "%", Location = new Point(245, gy+2), Width = 30 });

            gy += 35;
            grp.Controls.Add(new Label { Text = "Chi phí ko XĐ KL (TT):", Location = new Point(10, gy), Width = 160 });
            txtTT = new TextBox { Location = new Point(180, gy), Width = 60 };
            grp.Controls.Add(txtTT);
            grp.Controls.Add(new Label { Text = "%", Location = new Point(245, gy+2), Width = 30 });

            gy += 35;
            grp.Controls.Add(new Label { Text = "Lợi nhuận (TNCTTT):", Location = new Point(10, gy), Width = 160 });
            txtTNCTTT = new TextBox { Location = new Point(180, gy), Width = 60, Text = "5,5" }; // Mặc định chung
            grp.Controls.Add(txtTNCTTT);
            grp.Controls.Add(new Label { Text = "%", Location = new Point(245, gy+2), Width = 30 });

            gy += 35;
            grp.Controls.Add(new Label { Text = "Thuế GTGT:", Location = new Point(10, gy), Width = 160 });
            txtGTGT = new TextBox { Location = new Point(180, gy), Width = 60, Text = "8" };
            grp.Controls.Add(txtGTGT);
            grp.Controls.Add(new Label { Text = "%", Location = new Point(245, gy+2), Width = 30 });

            gy += 35;
            grp.Controls.Add(new Label { Text = "Nhà tạm (LT):", Location = new Point(10, gy), Width = 160 });
            txtNhaTam = new TextBox { Location = new Point(180, gy), Width = 60, Text = "1,1" };
            grp.Controls.Add(txtNhaTam);
            grp.Controls.Add(new Label { Text = "%", Location = new Point(245, gy+2), Width = 30 });

            y += 240;
            btnXuatExcel = new Button { Text = "Xuất 7 Bảng Biểu", Location = new Point(10, y), Size = new Size(320, 40), BackColor = Color.FromArgb(33, 115, 70), ForeColor = Color.White };
            btnXuatExcel.Click += BtnXuatExcel_Click;
            leftPanel.Controls.Add(btnXuatExcel);

            // Hook TextChanged to automatically preview
            txtCPC.TextChanged += (s, e) => BtnTinhToan_Click(null, null);
            txtTT.TextChanged += (s, e) => BtnTinhToan_Click(null, null);
            txtTNCTTT.TextChanged += (s, e) => BtnTinhToan_Click(null, null);
            txtGTGT.TextChanged += (s, e) => BtnTinhToan_Click(null, null);
            txtNhaTam.TextChanged += (s, e) => BtnTinhToan_Click(null, null);

            // Right Panel (Grid)
            dgvPreview = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White
            };
            dgvPreview.Columns.Add("DienGiai", "Nội dung chi phí");
            dgvPreview.Columns.Add("KyHieu", "Ký hiệu");
            dgvPreview.Columns.Add("CachTinh", "Cách tính");
            dgvPreview.Columns.Add("GiaTri", "Giá trị (đồng)");
            dgvPreview.Columns[0].FillWeight = 40;
            dgvPreview.Columns[1].FillWeight = 15;
            dgvPreview.Columns[2].FillWeight = 20;
            dgvPreview.Columns[3].FillWeight = 25;
            dgvPreview.Columns[3].DefaultCellStyle.Format = "N0";
            dgvPreview.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            
            dgvPreview.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            dgvPreview.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            rightPanel.Controls.Add(dgvPreview);
        }

        private void LoadLoaiCongTrinh()
        {
            var data = _cpcRepo.GetAll();
            var distinctTypes = data.Select(x => x.LoaiCongTrinh).Distinct().ToList();
            cboLoaiCongTrinh.DataSource = distinctTypes;
        }

        private void CboLoaiCongTrinh_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboLoaiCongTrinh.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinh.SelectedItem.ToString();
            
            var data = _cpcRepo.GetByLoaiCongTrinh(loaiCT);
            var phanLoai = data.Where(x => !string.IsNullOrEmpty(x.PhanLoaiPhu))
                               .Select(x => x.PhanLoaiPhu)
                               .Distinct()
                               .ToList();
            
            cboPhanLoaiPhu.Items.Clear();
            if (phanLoai.Count == 0)
            {
                cboPhanLoaiPhu.Items.Add("--- Không có ---");
                cboPhanLoaiPhu.Enabled = false;
            }
            else
            {
                cboPhanLoaiPhu.Enabled = true;
                cboPhanLoaiPhu.Items.Add("--- Mặc định ---");
                foreach (var item in phanLoai)
                    cboPhanLoaiPhu.Items.Add(item);
            }
                
            cboPhanLoaiPhu.SelectedIndex = 0;

            // Bảng 3.6: Định mức Thu nhập chịu thuế tính trước
            if (loaiCT == "Công nghiệp" || loaiCT == "Giao thông")
                txtTNCTTT.Text = "6,0";
            else
                txtTNCTTT.Text = "5,5";

            TuDongTraTiLe();
        }

        private void TuDongTraTiLe()
        {
            if (cboLoaiCongTrinh.SelectedItem == null) return;
            string loaiCT = cboLoaiCongTrinh.SelectedItem.ToString();
            string phanLoai = cboPhanLoaiPhu.SelectedIndex > 0 ? cboPhanLoaiPhu.SelectedItem.ToString() : null;
            
            decimal.TryParse(txtQuyMo.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal quyMo);

            // Tra CPC
            var listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
            if (listCPC.Count == 0 && phanLoai != null)
                listCPC = _cpcRepo.GetByLoaiCongTrinh(loaiCT, null); // Fallback

            if (listCPC.Count > 0)
            {
                decimal cpc = InterpolationHelper.NoiSuyTiLeCPC(listCPC, quyMo);
                txtCPC.Text = cpc.ToString("0.000").Replace('.', ',');
            }

            // Tra TT
            var tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, phanLoai);
            if (tt == null && phanLoai != null)
                tt = _ttRepo.GetByLoaiCongTrinh(loaiCT, null);
                
            if (tt != null)
            {
                txtTT.Text = tt.TiLe.ToString("0.000").Replace('.', ',');
            }
            
            BtnTinhToan_Click(null, null);
        }

        private void BtnTinhToan_Click(object sender, EventArgs e)
        {
            decimal.TryParse(txtCPC.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal cpc);
            decimal.TryParse(txtTT.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tt);
            decimal.TryParse(txtTNCTTT.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tncttt);
            decimal.TryParse(txtGTGT.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal gtgt);
            decimal.TryParse(txtNhaTam.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal nhatam);

            decimal tongMay = _tongT - _tongNC - (_duToan.DanhSachHangMuc.Sum(hm => hm.DanhSachCongTac.Sum(c => c.ThanhTienVL)));
            decimal tongVL = _tongT - _tongNC - tongMay;

            // Theo yêu cầu mới của người dùng: CPC tính trên T (tổng chi phí trực tiếp) thay vì NC
            var kq = _calcService.Tinh(tongVL, _tongNC, tongMay, cpc, tt, tncttt, gtgt, nhatam, "T");
            _duToan.ChiPhiXD = kq;

            dgvPreview.Rows.Clear();
            dgvPreview.Rows.Add("I. Chi phí trực tiếp", "T", "VL + NC + M", kq.T);
            dgvPreview.Rows.Add("- Chi phí vật liệu", "VL", "", kq.VL);
            dgvPreview.Rows.Add("- Chi phí nhân công", "NC", "", kq.NC);
            dgvPreview.Rows.Add("- Chi phí máy", "M", "", kq.M);
            dgvPreview.Rows.Add("II. Chi phí gián tiếp", "GT", "CPC + TT", kq.GT);
            dgvPreview.Rows.Add($"- Chi phí chung ({cpc}%)", "CPC", "T × tỷ lệ", kq.CPC);
            dgvPreview.Rows.Add($"- CP không xác định KL ({tt}%)", "TT", "T × tỷ lệ", kq.TT);
            dgvPreview.Rows.Add($"III. Thu nhập chịu thuế tính trước ({tncttt}%)", "TL", "(T + GT) × tỷ lệ", kq.TL);
            dgvPreview.Rows.Add("IV. Chi phí xây dựng trước thuế", "G", "T + GT + TL", kq.G);
            dgvPreview.Rows.Add($"V. Thuế GTGT ({gtgt}%)", "GTGT", "G × tỷ lệ", kq.GTGT);
            dgvPreview.Rows.Add("VI. Chi phí xây dựng sau thuế", "Gxd", "G + GTGT", kq.Gxd);
            dgvPreview.Rows.Add($"VII. Chi phí nhà tạm ({nhatam}%)", "LT", "Gxd × tỷ lệ", kq.LT);
            
            var row = new DataGridViewRow();
            row.CreateCells(dgvPreview, "VIII. TỔNG CỘNG CHI PHÍ XÂY DỰNG", "GXD", "Gxd + LT", kq.GXD);
            row.DefaultCellStyle.Font = new Font(dgvPreview.Font, FontStyle.Bold);
            row.DefaultCellStyle.BackColor = Color.LightYellow;
            dgvPreview.Rows.Add(row);
        }

        private void BtnXuatExcel_Click(object sender, EventArgs e)
        {
            if (_duToan.ChiPhiXD == null)
            {
                BtnTinhToan_Click(null, null); // Tính trước nếu chưa tính
            }

            try
            {
                var xuatService = new XuatBangBieuService();
                xuatService.Xuat7BangBieu(_duToan);
                MessageBox.Show("Xuất 7 bảng biểu thành công!", "AIE Dự Toán", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
