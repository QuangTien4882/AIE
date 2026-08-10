using System;

namespace AIE.Core.Models
{
    public class TuDienDongNghia
    {
        public int Id { get; set; }
        
        /// <summary>
        /// Tên nguyên bản (từ phần mềm Dự toán bên ngoài) cần chuẩn hóa, ví dụ: "Đá dăm 1x2"
        /// </summary>
        public string TenGoc { get; set; } = string.Empty;

        /// <summary>
        /// Tên chuẩn trong CSDL AIE (từ Sở Xây Dựng Đà Nẵng), ví dụ: "Đá 1x2"
        /// </summary>
        public string TenChuan { get; set; } = string.Empty;

        /// <summary>
        /// VL (Vật liệu), NC (Nhân công), MAY (Máy thi công)
        /// </summary>
        public string LoaiVatTu { get; set; } = string.Empty;
        
        public string NgayTao { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    }
}
