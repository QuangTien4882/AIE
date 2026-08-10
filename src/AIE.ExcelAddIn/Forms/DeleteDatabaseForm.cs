using System;
using System.Drawing;
using System.Windows.Forms;
using AIE.Data;

namespace AIE.ExcelAddIn.Forms
{
    public class DeleteDatabaseForm : Form
    {
        private CheckBox chkDinhMuc;
        private CheckBox chkVatLieu;
        private CheckBox chkNhanCong;
        private CheckBox chkMayThiCong;
        private Button btnDelete;
        private Button btnCancel;

        public DeleteDatabaseForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Xóa dữ liệu Database";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Width = 420;
            this.Height = 280;
            
            this.Font = new Font("Be Vietnam Pro", 9.5F);
            this.BackColor = Color.White;

            var lblTitle = new Label
            {
                Text = "Chọn các dữ liệu bạn muốn xóa khỏi Database:",
                AutoSize = true,
                Location = new Point(20, 20),
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F)
            };

            chkDinhMuc = new CheckBox { Text = "Định mức xây dựng", Location = new Point(40, 55), AutoSize = true, Checked = true };
            chkVatLieu = new CheckBox { Text = "Giá Vật liệu", Location = new Point(40, 85), AutoSize = true, Checked = true };
            chkNhanCong = new CheckBox { Text = "Giá Nhân công", Location = new Point(40, 115), AutoSize = true, Checked = true };
            chkMayThiCong = new CheckBox { Text = "Giá Máy thi công", Location = new Point(40, 145), AutoSize = true, Checked = true };

            btnDelete = new Button
            {
                Text = "Xóa dữ liệu",
                Location = new Point(80, 190),
                Width = 120,
                Height = 35,
                BackColor = Color.FromArgb(231, 76, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Be Vietnam Pro SemiBold", 9.5F)
            };
            btnDelete.FlatAppearance.BorderSize = 0;
            btnDelete.Click += BtnDelete_Click;

            btnCancel = new Button
            {
                Text = "Hủy bỏ",
                Location = new Point(210, 190),
                Width = 100,
                Height = 35,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, e) => this.Close();

            this.Controls.Add(lblTitle);
            this.Controls.Add(chkDinhMuc);
            this.Controls.Add(chkVatLieu);
            this.Controls.Add(chkNhanCong);
            this.Controls.Add(chkMayThiCong);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnCancel);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!chkDinhMuc.Checked && !chkVatLieu.Checked && !chkNhanCong.Checked && !chkMayThiCong.Checked)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một loại dữ liệu để xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show("Bạn có chắc chắn muốn xóa các dữ liệu đã chọn? Thao tác này không thể hoàn tác!", "Cảnh báo Xóa Database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                try
                {
                    var db = new DatabaseManager();
                    if (chkDinhMuc.Checked) db.ClearAllDinhMuc();
                    if (chkVatLieu.Checked) db.ClearAllVatLieu();
                    if (chkNhanCong.Checked) db.ClearAllNhanCong();
                    if (chkMayThiCong.Checked) db.ClearAllMayThiCong();

                    MessageBox.Show("Đã xóa dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa Database: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
