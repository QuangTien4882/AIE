using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.ExcelAddIn.Helpers;

namespace AIE.ExcelAddIn.Forms
{
    public class ChonBoDonGiaForm : Form
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        private BoDonGiaRepository _repo;
        private ListBox lstBoDonGia;
        private TextBox txtTenBoMoi;
        private ComboBox cboNguonCopy;
        private Button btnChon;
        private Button btnTaoMoi;

        public int SelectedBoDonGiaId { get; private set; }
        public string SelectedTenBo { get; private set; }

        public ChonBoDonGiaForm(BoDonGiaRepository repo)
        {
            _repo = repo;
            Text = "Chọn Bộ Đơn Giá";
            Size = new Size(400, 420);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            MinimizeBox = true;

            InitializeComponents();
            FormStateHelper.Attach(this);
            LoadData();
        }

        private void InitializeComponents()
        {
            var lblSelect = new Label { Text = "Chọn bộ đơn giá có sẵn:", Location = new Point(12, 12), AutoSize = true };
            lstBoDonGia = new ListBox { Location = new Point(12, 35), Size = new Size(360, 150) };
            lstBoDonGia.DisplayMember = "TenBo";
            lstBoDonGia.ValueMember = "Id";

            btnChon = new Button { Text = "Sử dụng bộ đã chọn", Location = new Point(12, 195), Size = new Size(180, 30) };
            btnChon.Click += BtnChon_Click;

            var btnDoiTen = new Button { Text = "Đổi tên", Location = new Point(200, 195), Size = new Size(80, 30) };
            btnDoiTen.Click += BtnDoiTen_Click;

            var btnXoa = new Button { Text = "Xóa", Location = new Point(288, 195), Size = new Size(84, 30), ForeColor = Color.Red };
            btnXoa.Click += BtnXoa_Click;

            var lblOr = new Label { Text = "HOẶC TẠO MỚI:", Location = new Point(12, 235), AutoSize = true, Font = new Font(Font, FontStyle.Bold) };
            
            var lblTenMoi = new Label { Text = "Tên bộ mới:", Location = new Point(12, 265), AutoSize = true };
            txtTenBoMoi = new TextBox { Location = new Point(100, 262), Size = new Size(272, 25) };
            SendMessage(txtTenBoMoi.Handle, 0x1501, 1, "VD: Quý 3-2026 Đà Nẵng");

            var lblNguon = new Label { Text = "Nguồn copy:", Location = new Point(12, 295), AutoSize = true };
            cboNguonCopy = new ComboBox { Location = new Point(100, 292), Size = new Size(272, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            
            btnTaoMoi = new Button { Text = "Tạo mới", Location = new Point(100, 325), Size = new Size(272, 30) };
            btnTaoMoi.Click += BtnTaoMoi_Click;

            Controls.Add(lblSelect);
            Controls.Add(lstBoDonGia);
            Controls.Add(btnChon);
            Controls.Add(btnDoiTen);
            Controls.Add(btnXoa);
            Controls.Add(lblOr);
            Controls.Add(lblTenMoi);
            Controls.Add(txtTenBoMoi);
            Controls.Add(lblNguon);
            Controls.Add(cboNguonCopy);
            Controls.Add(btnTaoMoi);
        }

        private void BtnDoiTen_Click(object sender, EventArgs e)
        {
            if (lstBoDonGia.SelectedItem is BoDonGiaRepository.BoDonGiaInfo info)
            {
                // Simple prompt using Microsoft.VisualBasic or a custom form. Since we don't have Microsoft.VisualBasic included, let's create a quick InputBox form inline.
                using (var inputForm = new Form() { Width = 400, Height = 150, Text = "Đổi tên bộ đơn giá", StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false })
                {
                    var lbl = new Label() { Left = 20, Top = 20, Text = "Nhập tên mới:", AutoSize = true };
                    var txt = new TextBox() { Left = 20, Top = 45, Width = 340, Text = info.TenBo };
                    var btnOk = new Button() { Text = "Đồng ý", Left = 180, Width = 80, Top = 75, DialogResult = DialogResult.OK };
                    var btnCancel = new Button() { Text = "Hủy", Left = 280, Width = 80, Top = 75, DialogResult = DialogResult.Cancel };
                    
                    inputForm.Controls.Add(lbl);
                    inputForm.Controls.Add(txt);
                    inputForm.Controls.Add(btnOk);
                    inputForm.Controls.Add(btnCancel);
                    inputForm.AcceptButton = btnOk;
                    inputForm.CancelButton = btnCancel;

                    if (inputForm.ShowDialog() == DialogResult.OK)
                    {
                        var newName = txt.Text.Trim();
                        if (!string.IsNullOrEmpty(newName) && newName != info.TenBo)
                        {
                            _repo.UpdateName(info.Id, newName);
                            LoadData(); // reload
                            
                            // re-select
                            for (int i = 0; i < lstBoDonGia.Items.Count; i++)
                            {
                                if (((BoDonGiaRepository.BoDonGiaInfo)lstBoDonGia.Items[i]).Id == info.Id)
                                {
                                    lstBoDonGia.SelectedIndex = i;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một bộ đơn giá để đổi tên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnXoa_Click(object sender, EventArgs e)
        {
            if (lstBoDonGia.SelectedItem is BoDonGiaRepository.BoDonGiaInfo info)
            {
                var result = MessageBox.Show(
                    $"Bạn có chắc chắn muốn xóa bộ đơn giá \"{info.TenBo}\"?\n\nTất cả dữ liệu giá trong bộ này sẽ bị xóa vĩnh viễn.",
                    "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        _repo.Delete(info.Id);
                        LoadData();
                        MessageBox.Show("Đã xóa bộ đơn giá thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi xóa: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một bộ đơn giá để xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LoadData()
        {
            var ds = _repo.GetAll().ToList();
            lstBoDonGia.DataSource = ds;
            
            // Setup source selection combo box
            var dsCopy = new System.Collections.Generic.List<BoDonGiaRepository.BoDonGiaInfo>();
            dsCopy.Add(new BoDonGiaRepository.BoDonGiaInfo { Id = 0, TenBo = "[ Không sao chép ]" });
            dsCopy.AddRange(ds);
            cboNguonCopy.DisplayMember = "TenBo";
            cboNguonCopy.ValueMember = "Id";
            cboNguonCopy.DataSource = dsCopy;

            if (ds.Count > 0)
            {
                lstBoDonGia.SelectedIndex = 0;
            }
        }

        private void BtnChon_Click(object sender, EventArgs e)
        {
            if (lstBoDonGia.SelectedItem is BoDonGiaRepository.BoDonGiaInfo info)
            {
                SelectedBoDonGiaId = info.Id;
                SelectedTenBo = info.TenBo;
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một bộ đơn giá.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnTaoMoi_Click(object sender, EventArgs e)
        {
            string tenBo = txtTenBoMoi.Text.Trim();
            if (string.IsNullOrEmpty(tenBo))
            {
                MessageBox.Show("Vui lòng nhập tên bộ đơn giá mới.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int newId = _repo.Create(tenBo);
            
            // Sao chép dữ liệu từ nguồn đã chọn
            if (cboNguonCopy.SelectedItem is BoDonGiaRepository.BoDonGiaInfo info && info.Id > 0)
            {
                _repo.CopyPrices(info.Id, newId);
            }

            SelectedBoDonGiaId = newId;
            SelectedTenBo = tenBo;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
