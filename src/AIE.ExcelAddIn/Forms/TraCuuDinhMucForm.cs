using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.Core.Models;

namespace AIE.ExcelAddIn.Forms
{
    public class TraCuuDinhMucForm : Form
    {
        private TextBox txtSearch;
        private Button btnSearch;
        private Button btnPhuLuc;
        private ContextMenuStrip menuPhuLuc;
        private DataGridView dgvCongTac;
        private DataGridView dgvHaoPhi;
        private DatabaseManager _dbManager;
        private CongTacRepository _repo;
        private System.Windows.Forms.Timer _searchTimer;
        private BindingSource _bsCongTac;
        private List<CongTacXayDung> _allCongTac;
        
        private string _settingsPath;

        public TraCuuDinhMucForm()
        {
            _dbManager = new DatabaseManager();
            _repo = new CongTacRepository(_dbManager.Context);
            _bsCongTac = new BindingSource();
            
            string appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
            if (!Directory.Exists(appData)) Directory.CreateDirectory(appData);
            _settingsPath = Path.Combine(appData, "grid_settings.txt");

            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Tra Cứu Định Mức (Thông tư 38)";
            
            var workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Size = new Size((int)(workingArea.Width * 0.9), (int)(workingArea.Height * 0.9));
            this.StartPosition = FormStartPosition.CenterScreen;
            
            this.Font = new Font("Be Vietnam Pro", 9.5F);
            this.BackColor = Color.FromArgb(245, 246, 250);
            this.ShowIcon = false;

            // --- HEADER PANEL ---
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.White };
            pnlTop.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, pnlTop.ClientRectangle, Color.FromArgb(220, 221, 225), ButtonBorderStyle.Solid);
            };
            
            var lblPhuLuc = new Label { 
                Text = "Bộ định mức:", 
                AutoSize = true, 
                Location = new Point(20, 20),
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F),
                ForeColor = Color.FromArgb(47, 54, 64)
            };

            btnPhuLuc = new Button {
                Text = "Chọn Phụ Lục ▼",
                Location = new Point(130, 15),
                Width = 200,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                Cursor = Cursors.Hand
            };
            btnPhuLuc.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);

            menuPhuLuc = new ContextMenuStrip();
            string[] phuLucs = { 
                "Phụ lục I - Khảo sát (Mã C)", 
                "Phụ lục II - Xây dựng (Mã A)", 
                "Phụ lục III - Lắp đặt hệ thống (Mã B)", 
                "Phụ lục IV - Lắp đặt máy và TB (Mã M)", 
                "Phụ lục V - Thí nghiệm (Mã D)", 
                "Phụ lục VI - Sửa chữa (Mã S)" 
            };
            
            foreach (var pl in phuLucs)
            {
                var item = new ToolStripMenuItem(pl) { CheckOnClick = true, Checked = true };
                item.CheckedChanged += (s, e) => FilterData();
                menuPhuLuc.Items.Add(item);
            }
            
            menuPhuLuc.Items.Add(new ToolStripSeparator());
            var itemBoChon = new ToolStripMenuItem("Bỏ chọn toàn bộ (Clear all)");
            itemBoChon.Click += (s, e) => {
                foreach (ToolStripItem it in menuPhuLuc.Items)
                {
                    if (it is ToolStripMenuItem tsi && tsi != itemBoChon)
                    {
                        tsi.Checked = false;
                    }
                }
                FilterData();
            };
            menuPhuLuc.Items.Add(itemBoChon);
            
            btnPhuLuc.Click += (s, e) => menuPhuLuc.Show(btnPhuLuc, new Point(0, btnPhuLuc.Height));
            
            var lblSearch = new Label { 
                Text = "Tìm kiếm:", 
                AutoSize = true, 
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F),
                ForeColor = Color.FromArgb(47, 54, 64)
            };
            
            txtSearch = new TextBox { 
                Width = 300,
                Font = new Font("Be Vietnam Pro", 9.5F),
                BorderStyle = BorderStyle.FixedSingle
            };
            
            _searchTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _searchTimer.Tick += (s, e) => {
                _searchTimer.Stop();
                FilterData();
            };

            txtSearch.TextChanged += (s, e) => {
                _searchTimer.Stop();
                _searchTimer.Start();
            };
            
            btnSearch = new Button { 
                Text = "🔍", 
                Width = 40,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 168, 255),
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI Emoji", 10F)
            };
            btnSearch.FlatAppearance.BorderSize = 0;
            ToolTip tt = new ToolTip();
            tt.SetToolTip(btnSearch, "Tìm kiếm định mức");
            btnSearch.Click += (s, e) => FilterData();

            pnlTop.Controls.Add(lblPhuLuc);
            pnlTop.Controls.Add(btnPhuLuc);
            pnlTop.Controls.Add(lblSearch);
            pnlTop.Controls.Add(txtSearch);
            pnlTop.Controls.Add(btnSearch);

            // Tính vị trí search controls bám sát bên phải
            Action layoutSearch = () => {
                int right = pnlTop.ClientSize.Width - 15;
                btnSearch.Location = new Point(right - btnSearch.Width, 16);
                txtSearch.Location = new Point(btnSearch.Left - txtSearch.Width - 5, 18);
                lblSearch.Location = new Point(txtSearch.Left - lblSearch.Width - 5, 20);
            };
            pnlTop.Resize += (s, e) => layoutSearch();
            pnlTop.Layout += (s, e) => layoutSearch();

            // --- SPLIT CONTAINER ---
            var splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.None,
                SplitterDistance = (int)((workingArea.Height * 0.9 - 60) * 0.06),
                SplitterWidth = 6,
                BackColor = Color.FromArgb(200, 200, 200)
            };

            // DGV Công tác
            dgvCongTac = CreateModernGrid();
            dgvCongTac.SelectionChanged += DgvCongTac_SelectionChanged;
            SetupCongTacColumns();
            
            var pnlGrid1 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = Color.FromArgb(245, 246, 250) };
            pnlGrid1.Controls.Add(dgvCongTac);
            splitContainer.Panel1.Controls.Add(pnlGrid1);

            // DGV Hao phí
            dgvHaoPhi = CreateModernGrid();
            SetupHaoPhiColumns();
            
            var pnlGrid2 = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10), BackColor = Color.FromArgb(245, 246, 250) };
            pnlGrid2.Controls.Add(dgvHaoPhi);
            splitContainer.Panel2.Controls.Add(pnlGrid2);

            this.Controls.Add(splitContainer);
            this.Controls.Add(pnlTop);

            this.Load += TraCuuDinhMucForm_Load;
            this.FormClosing += TraCuuDinhMucForm_FormClosing;
        }

        private DataGridView CreateModernGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                GridColor = Color.FromArgb(200, 200, 200),
                EnableHeadersVisualStyles = false,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                ColumnHeadersHeight = 35,
                RowTemplate = { Height = 30 },
                DefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Be Vietnam Pro", 9F) }
            };
        }

        private void SetupCongTacColumns()
        {
            dgvCongTac.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgvCongTac.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvCongTac.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro SemiBold", 9.5F);
            dgvCongTac.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            // Xám ghi selection
            dgvCongTac.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 215, 220);
            dgvCongTac.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieu", DataPropertyName = "MaHieu", HeaderText = "Mã hiệu", Width = 100 });
            dgvCongTac.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenCongTac", DataPropertyName = "TenCongTac", HeaderText = "Tên công tác xây dựng", Width = 500 });
            var colDonVi = new DataGridViewTextBoxColumn { Name = "DonVi", DataPropertyName = "DonVi", HeaderText = "Đơn vị", Width = 100 };
            colDonVi.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvCongTac.Columns.Add(colDonVi);
        }

        private void SetupHaoPhiColumns()
        {
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.Font = new Font("Be Vietnam Pro SemiBold", 9.5F);
            dgvHaoPhi.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            // Xám ghi selection
            dgvHaoPhi.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 215, 220);
            dgvHaoPhi.DefaultCellStyle.SelectionForeColor = Color.Black;

            var colLoai = new DataGridViewTextBoxColumn { Name = "LoaiHaoPhiStr", DataPropertyName = "LoaiHaoPhiStr", HeaderText = "Loại hao phí", Width = 120 };
            colLoai.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvHaoPhi.Columns.Add(colLoai);
            
            dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaHieuHP", DataPropertyName = "MaHieuHP", HeaderText = "Mã vật tư", Width = 100 });
            dgvHaoPhi.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenHaoPhi", DataPropertyName = "TenHaoPhi", HeaderText = "Tên vật tư / Nhân công / Máy TC", Width = 500 });
            
            var colDonVi = new DataGridViewTextBoxColumn { Name = "DonVi", DataPropertyName = "DonVi", HeaderText = "Đơn vị", Width = 100 };
            colDonVi.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvHaoPhi.Columns.Add(colDonVi);
            
            var colDinhMuc = new DataGridViewTextBoxColumn { Name = "DinhMuc", DataPropertyName = "DinhMuc", HeaderText = "Định mức hao phí", Width = 150 };
            colDinhMuc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            colDinhMuc.DefaultCellStyle.Format = "N4";
            dgvHaoPhi.Columns.Add(colDinhMuc);
        }

        private void TraCuuDinhMucForm_Load(object sender, EventArgs e)
        {
            try
            {
                _allCongTac = _repo.GetAllCongTac().OrderBy(x => x.MaHieu).ToList();
                _bsCongTac.DataSource = _allCongTac;
                dgvCongTac.DataSource = _bsCongTac;
                
                LoadColumnSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FilterData()
        {
            var keyword = txtSearch.Text.Trim().ToLower();
            
            var allowedPrefixes = new List<string>();
            if (((ToolStripMenuItem)menuPhuLuc.Items[0]).Checked) allowedPrefixes.Add("c");
            if (((ToolStripMenuItem)menuPhuLuc.Items[1]).Checked) allowedPrefixes.Add("a");
            if (((ToolStripMenuItem)menuPhuLuc.Items[2]).Checked) allowedPrefixes.Add("b");
            if (((ToolStripMenuItem)menuPhuLuc.Items[3]).Checked) allowedPrefixes.Add("m");
            if (((ToolStripMenuItem)menuPhuLuc.Items[4]).Checked) allowedPrefixes.Add("d");
            if (((ToolStripMenuItem)menuPhuLuc.Items[5]).Checked) allowedPrefixes.Add("s");

            var filtered = _allCongTac.Where(x => 
                (x.MaHieu != null && allowedPrefixes.Any(p => x.MaHieu.ToLower().StartsWith(p))) &&
                (string.IsNullOrEmpty(keyword) || 
                 (x.MaHieu != null && x.MaHieu.ToLower().Contains(keyword)) || 
                 (x.TenCongTac != null && x.TenCongTac.ToLower().Contains(keyword)))
            ).ToList();

            _bsCongTac.DataSource = filtered;
        }

        private void DgvCongTac_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvCongTac.SelectedRows.Count > 0)
            {
                var ct = dgvCongTac.SelectedRows[0].DataBoundItem as CongTacXayDung;
                if (ct != null && ct.DanhSachHaoPhi != null)
                {
                    var viewList = ct.DanhSachHaoPhi.Select(hp => new {
                        LoaiHaoPhiStr = hp.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.VL ? "VL" : 
                                       (hp.LoaiHaoPhi == AIE.Core.Enums.LoaiHaoPhi.NC ? "NC" : "MTC"),
                        hp.MaHieuHP,
                        hp.TenHaoPhi,
                        hp.DonVi,
                        hp.DinhMuc
                    }).ToList();
                    
                    dgvHaoPhi.DataSource = viewList;
                }
                else
                {
                    dgvHaoPhi.DataSource = null;
                }
            }
            else
            {
                dgvHaoPhi.DataSource = null;
            }
        }

        private void TraCuuDinhMucForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveColumnSettings();
        }

        // --- SAVE / LOAD COLUMN WIDTHS ---
        private void SaveColumnSettings()
        {
            try
            {
                var lines = new List<string>();
                foreach (DataGridViewColumn col in dgvCongTac.Columns)
                    if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.NotSet || col.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                        lines.Add($"C|{col.Name}|{col.Width}");

                foreach (DataGridViewColumn col in dgvHaoPhi.Columns)
                    if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.NotSet || col.AutoSizeMode == DataGridViewAutoSizeColumnMode.None)
                        lines.Add($"H|{col.Name}|{col.Width}");

                File.WriteAllLines(_settingsPath, lines);
            }
            catch { /* ignore */ }
        }

        private void LoadColumnSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var lines = File.ReadAllLines(_settingsPath);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 3 && int.TryParse(parts[2], out int width))
                        {
                            if (parts[0] == "C" && dgvCongTac.Columns[parts[1]] != null)
                                dgvCongTac.Columns[parts[1]].Width = width;
                            else if (parts[0] == "H" && dgvHaoPhi.Columns[parts[1]] != null)
                                dgvHaoPhi.Columns[parts[1]].Width = width;
                        }
                    }
                }
            }
            catch { /* ignore */ }
        }
    }
}
