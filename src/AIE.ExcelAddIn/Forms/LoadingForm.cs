using System;
using System.Drawing;
using System.Windows.Forms;

namespace AIE.ExcelAddIn.Forms
{
    public class LoadingForm : Form
    {
        public LoadingForm(string message = "Đang xử lý, vui lòng đợi...")
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(300, 80);
            this.BackColor = Color.White;
            this.TopMost = true;

            var lbl = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 12f, FontStyle.Regular),
                ForeColor = Color.FromArgb(0, 120, 215)
            };

            this.Controls.Add(lbl);

            // Thêm viền
            this.Paint += (s, e) =>
            {
                ControlPaint.DrawBorder(e.Graphics, this.ClientRectangle, Color.FromArgb(0, 120, 215), ButtonBorderStyle.Solid);
            };
        }
    }
}
