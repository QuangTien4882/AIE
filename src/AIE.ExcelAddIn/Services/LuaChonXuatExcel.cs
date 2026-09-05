using System.Collections.Generic;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Các tùy chọn bảng biểu cần xuất sang Excel.
    /// </summary>
    public class LuaChonXuatExcel
    {
        // 1. Nhóm Bảng Tổng hợp kinh phí (TT 36/2026/TT-BXD)
        public bool XuatTongMucDauTu { get; set; } = true;         // sheet: TongMucDauTu (Bảng 1.2)
        public bool XuatTongHopDuToan { get; set; } = false;        // sheet: TH_DuToan (Bảng 2.1)
        public bool XuatChiPhiXayDung { get; set; } = true;         // sheet: TH_ChiPhiXD (Bảng 3.8)

        // 2. Nhóm Bảng biểu Kỹ thuật & Dự toán chi tiết
        public bool XuatDuToanChiTiet { get; set; } = true;         // sheet: DuToan
        public bool XuatPhanTichDonGia { get; set; } = true;        // sheet: PhanTich_DonGia
        public bool XuatTongHopVatLieu { get; set; } = true;        // sheet: TH_VatLieu
        public bool XuatTongHopNhanCong { get; set; } = true;       // sheet: TH_NhanCong
        public bool XuatTongHopCaMay { get; set; } = true;          // sheet: TH_CaMay
        public bool XuatChietTinhCuocVC { get; set; } = true;       // sheet: ChietTinh_CuocVC
        public bool XuatHeSoDieuChinh { get; set; } = true;         // sheet: HeSo_DieuChinh

        public List<string> GetDanhSachDaChon()
        {
            var list = new List<string>();
            if (XuatTongMucDauTu) list.Add("Bảng 1.2: Tổng mức đầu tư xây dựng (sheet 'TongMucDauTu')");
            if (XuatTongHopDuToan) list.Add("Bảng 2.1: Tổng hợp dự toán công trình (sheet 'TH_DuToan')");
            if (XuatChiPhiXayDung) list.Add("Bảng 3.8: Bảng tổng hợp chi phí xây dựng (sheet 'TH_ChiPhiXD')");
            if (XuatDuToanChiTiet) list.Add("Dự toán chi phí xây dựng công trình (sheet 'DuToan')");
            if (XuatPhanTichDonGia) list.Add("Bảng phân tích đơn giá chi tiết (sheet 'PhanTich_DonGia')");
            if (XuatTongHopVatLieu) list.Add("Bảng tổng hợp chênh lệch giá vật liệu (sheet 'TH_VatLieu')");
            if (XuatTongHopNhanCong) list.Add("Bảng tổng hợp & chênh lệch nhân công (sheet 'TH_NhanCong')");
            if (XuatTongHopCaMay) list.Add("Bảng tổng hợp & chênh lệch máy thi công (sheet 'TH_CaMay')");
            if (XuatChietTinhCuocVC) list.Add("Bảng chiết tính cước vận chuyển (sheet 'ChietTinh_CuocVC')");
            if (XuatHeSoDieuChinh) list.Add("Bảng xác định hệ số điều chỉnh (sheet 'HeSo_DieuChinh')");
            return list;
        }

        public bool CoItNhatMotBangDuocChon()
        {
            return XuatTongMucDauTu || XuatTongHopDuToan || XuatChiPhiXayDung ||
                   XuatDuToanChiTiet || XuatPhanTichDonGia || XuatTongHopVatLieu ||
                   XuatTongHopNhanCong || XuatTongHopCaMay || XuatChietTinhCuocVC || XuatHeSoDieuChinh;
        }
    }
}
