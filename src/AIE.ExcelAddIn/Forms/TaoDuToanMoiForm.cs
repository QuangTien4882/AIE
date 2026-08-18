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
        private Button btnTao;
        private Button btnHuy;

        public TaoDuToanMoiForm()

        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Tạo Dự toán mới";
            this.Size = new Size(600, 220);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Be Vietnam Pro", 9.5f);
            this.BackColor = Color.White;

            int y = 20;
            int lblWidth = 120;
            int txtWidth = 430;

            // Tên công trình
            var lblTen = new Label { Text = "Tên công trình:", Location = new Point(20, y + 4), Size = new Size(lblWidth, 25) };
            txtTenCongTrinh = new TextBox { Location = new Point(140, y), Size = new Size(txtWidth, 25) };
            y += 40;

            // Địa điểm
            var lblDiaDiem = new Label { Text = "Địa điểm:", Location = new Point(20, y + 4), Size = new Size(lblWidth, 25) };
            txtDiaDiem = new TextBox { Location = new Point(140, y), Size = new Size(txtWidth, 25) };
            txtDiaDiem.Text = "Đà Nẵng"; // Default
            y += 40;

            // Buttons
            y += 10;
            btnTao = new Button { Text = "Khởi tạo", Location = new Point(190, y), Size = new Size(100, 35), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnTao.FlatAppearance.BorderSize = 0;
            btnTao.Click += BtnTao_Click;

            btnHuy = new Button { Text = "Hủy", Location = new Point(310, y), Size = new Size(100, 35), FlatStyle = FlatStyle.Flat };
            btnHuy.FlatAppearance.BorderColor = Color.LightGray;
            btnHuy.Click += (s, e) => this.Close();

            this.Controls.Add(lblTen);
            this.Controls.Add(txtTenCongTrinh);
            this.Controls.Add(lblDiaDiem);
            this.Controls.Add(txtDiaDiem);
            this.Controls.Add(btnTao);
            this.Controls.Add(btnHuy);
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

            // --- COLUMN HEADERS ---
            ws.Cells[4, 1] = "STT";
            ws.Cells[4, 2] = "Mã hiệu";
            ws.Cells[4, 3] = "Tên công tác";
            ws.Cells[4, 4] = "Đơn vị";
            ws.Cells[4, 5] = "Khối lượng";
            ws.Cells[4, 6] = "Đơn giá";
            ws.Cells[5, 6] = "Vật liệu";
            ws.Cells[5, 7] = "Nhân công";
            ws.Cells[5, 8] = "Máy thi công";
            ws.Cells[4, 9] = "Thành tiền";

            ws.Range["A4:A5"].Merge();
            ws.Range["B4:B5"].Merge();
            ws.Range["C4:C5"].Merge();
            ws.Range["D4:D5"].Merge();
            ws.Range["E4:E5"].Merge();
            ws.Range["F4:H4"].Merge();
            ws.Range["I4:I5"].Merge();

            Excel.Range columns = ws.Range[ws.Cells[4, 1], ws.Cells[5, 9]];
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
            ws.Cells[6, 2].Select();
            app.ActiveWindow.FreezePanes = false;
            app.ActiveWindow.SplitRow = 5;
            app.ActiveWindow.SplitColumn = 0;
            app.ActiveWindow.FreezePanes = true;
        }
    }
}
