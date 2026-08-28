using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using AIE.Core.Enums;
using System.Linq;
using Button = System.Windows.Forms.Button;
using TextBox = System.Windows.Forms.TextBox;
using Label = System.Windows.Forms.Label;
using ComboBox = System.Windows.Forms.ComboBox;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;
using Excel = Microsoft.Office.Interop.Excel;

namespace AIE.ExcelAddIn.Forms
{
    public class TaoDuToanMoiForm : Form
    {
        private TextBox txtTenCongTrinh;
        private TextBox txtDiaDiem;
        private ComboBox cboVung;
        private Button btnTao;
        private Button btnHuy;

        public TaoDuToanMoiForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Tạo Dự toán mới";
            this.Size = new Size(620, 400);
            this.MinimumSize = new Size(580, 350);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Be Vietnam Pro", 9.5f);
            this.BackColor = Color.White;

            int y = 20;
            int lblWidth = 120;
            int txtWidth = 430;

            // Tên công trình
            var lblTen = new Label { Text = "Tên công trình:", Location = new Point(20, y + 4), Size = new Size(lblWidth, 25) };
            txtTenCongTrinh = new TextBox { Location = new Point(140, y), Size = new Size(txtWidth, 25), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            y += 40;

            // Địa điểm
            var lblDiaDiem = new Label { Text = "Địa điểm:", Location = new Point(20, y + 4), Size = new Size(lblWidth, 25) };
            txtDiaDiem = new TextBox { Location = new Point(140, y), Size = new Size(txtWidth, 25), Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            txtDiaDiem.Text = "Đà Nẵng"; // Default
            y += 40;

            // Vùng
            var lblVung = new Label { Text = "Vùng áp dụng:", Location = new Point(20, y + 4), Size = new Size(lblWidth, 25) };
            cboVung = new ComboBox { Location = new Point(140, y), Size = new Size(txtWidth, 25), DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            
            var vungList = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<Vung, string>>
            {
                new System.Collections.Generic.KeyValuePair<Vung, string>(Vung.VungII, "Vùng II"),
                new System.Collections.Generic.KeyValuePair<Vung, string>(Vung.VungIII, "Vùng III"),
                new System.Collections.Generic.KeyValuePair<Vung, string>(Vung.VungIV, "Vùng IV"),
                new System.Collections.Generic.KeyValuePair<Vung, string>(Vung.CuLaoCham, "Cù Lao Chàm")
            };
            cboVung.DisplayMember = "Value";
            cboVung.ValueMember = "Key";
            cboVung.DataSource = vungList;
            y += 35;

            // Scrollable description panel that fills remaining area
            var descPanel = new Panel { Location = new Point(140, y), Size = new Size(txtWidth, 120), AutoScroll = true, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom, BorderStyle = BorderStyle.None };
            var lblVungDesc = new Label { Location = new Point(0, 0), AutoSize = true, MaximumSize = new Size(txtWidth - 20, 0), Font = new Font("Be Vietnam Pro", 8.5f, FontStyle.Italic), ForeColor = Color.DimGray };
            descPanel.Controls.Add(lblVungDesc);
            
            // Set initial description text explicitly since setting SelectedIndex before Handle is created throws exceptions
            lblVungDesc.Text = "Gồm các phường: Hải Châu, Hòa Cường, Thanh Khê, An Khê, An Hải, Sơn Trà, Ngũ Hành Sơn, Hòa Khánh, Hải Vân, Liên Chiểu, Cẩm Lệ, Hòa Xuân, Tam Kỳ, Quảng Phú, Hương Trà, Bàn Thạch, Hội An, Hội An Đông, Hội An Tây và các xã: Hòa Vang, Hòa Tiến, Bà Nà.";

            cboVung.SelectedIndexChanged += (s, e) =>
            {
                if (cboVung.SelectedItem == null) return;
                var selectedVung = ((System.Collections.Generic.KeyValuePair<Vung, string>)cboVung.SelectedItem).Key;
                switch (selectedVung)
                {
                    case Vung.VungII: lblVungDesc.Text = "Gồm các phường: Hải Châu, Hòa Cường, Thanh Khê, An Khê, An Hải, Sơn Trà, Ngũ Hành Sơn, Hòa Khánh, Hải Vân, Liên Chiểu, Cẩm Lệ, Hòa Xuân, Tam Kỳ, Quảng Phú, Hương Trà, Bàn Thạch, Hội An, Hội An Đông, Hội An Tây và các xã: Hòa Vang, Hòa Tiến, Bà Nà."; break;
                    case Vung.VungIII: lblVungDesc.Text = "Gồm các phường: Điện Bàn, Điện Bàn Đông, An Thắng, Điện Bàn Bắc và các xã: Núi Thành, Tam Mỹ, Tam Anh, Đức Phú, Tam Xuân, Tam Hải, Tây Hồ, Chiên Đàn, Phú Ninh, Thăng Bình, Thăng An, Thăng Trường, Thăng Điền, Thăng Phú, Đồng Dương, Quế Sơn Trung, Quế Sơn, Xuân Phú, Nông Sơn, Quế Phước, Duy Nghĩa, Nam Phước, Duy Xuyên, Thu Bồn, Điện Bàn Tây, Gò Nổi, Đại Lộc, Hà Nha, Thượng Đức, Vu Gia, Phú Thuận."; break;
                    case Vung.VungIV: lblVungDesc.Text = "Gồm các xã: Lãnh Ngọc, Tiên Phước, Xã Thạnh Bình, Sơn Cẩm Hà, Trà Liên, Trà Giáp, Trà Tân, Trà Đốc, Trà My, Nam Trà My, Trà Tập, Trà Vân, Trà Linh, Trà Leng, Thạnh Mỹ, Bến Giằng, Nam Giang, Đắc Pring, La Dêê, La Êê, Sông Vàng, Sông Kôn, Đông Giang, Bến Hiên, Avương, Tây Giang, Hùng Sơn, Hiệp Đức, Việt An, Phước Trà, Khâm Đức, Phước Năng, Phước Chánh, Phước Thành, Phước Hiệp."; break;
                    case Vung.CuLaoCham: lblVungDesc.Text = "Khu vực xã đảo Tân Hiệp (Cù Lao Chàm)"; break;
                }
            };

            // Bottom panel for buttons
            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 55, BackColor = Color.White };
            btnTao = new Button { Text = "Khởi tạo", Size = new Size(100, 35), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnTao.FlatAppearance.BorderSize = 0;
            btnTao.Click += BtnTao_Click;

            btnHuy = new Button { Text = "Hủy", Size = new Size(100, 35), FlatStyle = FlatStyle.Flat };
            btnHuy.FlatAppearance.BorderColor = Color.LightGray;
            btnHuy.Click += (s, e) => this.Close();

            bottomPanel.Controls.Add(btnTao);
            bottomPanel.Controls.Add(btnHuy);

            // Center buttons in bottom panel
            bottomPanel.Resize += (s, e) => {
                int totalW = btnTao.Width + 20 + btnHuy.Width;
                int startX = (bottomPanel.Width - totalW) / 2;
                btnTao.Location = new Point(startX, 10);
                btnHuy.Location = new Point(startX + btnTao.Width + 20, 10);
            };

            this.Controls.Add(lblTen);
            this.Controls.Add(txtTenCongTrinh);
            this.Controls.Add(lblDiaDiem);
            this.Controls.Add(txtDiaDiem);
            this.Controls.Add(lblVung);
            this.Controls.Add(cboVung);
            this.Controls.Add(descPanel);
            this.Controls.Add(bottomPanel);
        }



        private void BtnTao_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTenCongTrinh.Text))
            {
                MessageBox.Show("Vui lòng nhập tên công trình!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                CreateExcelTemplate();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi khởi tạo file Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateExcelTemplate()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            
            // Create a new workbook if none is active
            Excel.Workbook wb = app.ActiveWorkbook;
            if (wb == null)
            {
                wb = app.Workbooks.Add();
            }

            Excel.Worksheet ws = null;
            var activeWs = wb.ActiveSheet as Excel.Worksheet;
            if (activeWs != null && activeWs.Name.StartsWith("Sheet"))
            {
                // Simple check if empty
                var range = activeWs.UsedRange;
                if (range.Rows.Count <= 1 && range.Columns.Count <= 1 && string.IsNullOrEmpty(activeWs.Cells[1, 1].Text))
                {
                    ws = activeWs;
                }
            }

            if (ws == null)
            {
                ws = wb.Worksheets.Add();
            }
            
            ws.Name = "DuToan_" + DateTime.Now.ToString("HHmmss");

            // Set entire sheet font to Times New Roman, 12
            ws.Cells.Font.Name = "Times New Roman";
            ws.Cells.Font.Size = 12;

            // --- HEADER INFO ---
            // Row 1: Tên dự án
            Excel.Range r1 = ws.Range["A1", "I1"];
            r1.Merge();
            r1.Value2 = ("TÊN DỰ ÁN: " + txtTenCongTrinh.Text).ToUpper();
            r1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            r1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            r1.Font.Bold = true;

            // Row 2: Địa điểm
            Excel.Range r2 = ws.Range["A2", "I2"];
            r2.Merge();
            r2.Value2 = ("ĐỊA ĐIỂM: " + txtDiaDiem.Text).ToUpper();
            r2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            r2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            r2.Font.Bold = true;

            // Lưu giá trị Vung enum vào ô ẩn J1
            var selectedVung = ((System.Collections.Generic.KeyValuePair<Vung, string>)cboVung.SelectedItem).Key;
            ws.Cells[1, 10] = (int)selectedVung;
            ((Excel.Range)ws.Cells[1, 10]).Font.Color = ColorTranslator.ToOle(Color.White);
            ((Excel.Range)ws.Columns[10]).ColumnWidth = 0; // Ẩn cột J

            // --- COLUMN HEADERS (Row 3-4) ---
            ws.Cells[3, 1] = "STT";
            ws.Cells[3, 2] = "Mã hiệu";
            ws.Cells[3, 3] = "Tên công tác";
            ws.Cells[3, 4] = "Đơn vị";
            ws.Cells[3, 5] = "Khối lượng";
            ws.Cells[3, 6] = "Đơn giá";
            ws.Cells[4, 6] = "Vật liệu";
            ws.Cells[4, 7] = "Nhân công";
            ws.Cells[4, 8] = "Máy thi công";
            ws.Cells[3, 9] = "Thành tiền";

            ws.Range["A3:A4"].Merge();
            ws.Range["B3:B4"].Merge();
            ws.Range["C3:C4"].Merge();
            ws.Range["D3:D4"].Merge();
            ws.Range["E3:E4"].Merge();
            ws.Range["F3:H3"].Merge();
            ws.Range["I3:I4"].Merge();

            Excel.Range columns = ws.Range[ws.Cells[3, 1], ws.Cells[4, 9]];
            columns.Font.Bold = true;
            columns.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(200, 220, 240));
            columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            columns.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Column widths
            ((Excel.Range)ws.Columns[1]).ColumnWidth = 5;
            ((Excel.Range)ws.Columns[2]).ColumnWidth = 12;
            ((Excel.Range)ws.Columns[3]).ColumnWidth = 45;
            ((Excel.Range)ws.Columns[4]).ColumnWidth = 8;
            ((Excel.Range)ws.Columns[5]).ColumnWidth = 12;
            ((Excel.Range)ws.Columns[6]).ColumnWidth = 15;
            ((Excel.Range)ws.Columns[7]).ColumnWidth = 15;
            ((Excel.Range)ws.Columns[8]).ColumnWidth = 15;
            ((Excel.Range)ws.Columns[9]).ColumnWidth = 18;

            // Formatting columns for numbers
            ((Excel.Range)ws.Columns[5]).NumberFormat = "#,##0.00";
            ((Excel.Range)ws.Columns[6]).NumberFormat = "#,##0";
            ((Excel.Range)ws.Columns[7]).NumberFormat = "#,##0";
            ((Excel.Range)ws.Columns[8]).NumberFormat = "#,##0";
            ((Excel.Range)ws.Columns[9]).NumberFormat = "#,##0";

            // Alignment rules
            // STT (1), Đơn vị (4): Center
            ((Excel.Range)ws.Columns[1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)ws.Columns[4]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // Tên công tác (3): Justify (or Left with wrap)
            ((Excel.Range)ws.Columns[3]).HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify;
            ((Excel.Range)ws.Columns[3]).WrapText = true;
            // Số lượng, giá, Thành tiền (5 -> 9): Right
            Excel.Range numberCols = ws.Range[ws.Columns[5], ws.Columns[9]];
            numberCols.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            
            // Vertical center for all
            ws.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            // Re-apply header horizontal alignment just in case column alignments overrode it
            columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Select the first data cell and freeze panes
            ws.Cells[5, 2].Select();
            app.ActiveWindow.FreezePanes = false;
            app.ActiveWindow.SplitRow = 4;
            app.ActiveWindow.SplitColumn = 0;
            app.ActiveWindow.FreezePanes = true;
        }
    }
}
