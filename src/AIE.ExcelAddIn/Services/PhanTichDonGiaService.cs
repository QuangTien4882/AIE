using AIE.Core.Models;
using AIE.Core.Enums;
using AIE.Data.Repositories;
using System.Linq;

namespace AIE.ExcelAddIn.Services;

public class PhanTichDonGiaService
{
    private readonly CongTacRepository _congTacRepo;

    public PhanTichDonGiaService(CongTacRepository congTacRepo)
    {
        _congTacRepo = congTacRepo;
    }

    public void TinhDonGiaChiTiet(DuToan duToan)
    {
        var bangTongHop = duToan.BangTongHop;
        
        // Tạo dictionary để tra cứu nhanh giá hiện trường
        var dictVL = bangTongHop.DanhSachVatLieu.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);
        var dictNC = bangTongHop.DanhSachNhanCong.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);
        var dictMay = bangTongHop.DanhSachMay.ToDictionary(x => x.MaVatTu, x => x.GiaHienTruong);

        foreach (var hm in duToan.DanhSachHangMuc)
        {
            foreach (var ct in hm.DanhSachCongTac)
            {
                var congTacChuan = _congTacRepo.GetByMaHieu(ct.MaHieu);
                if (congTacChuan == null) continue;
                
                ct.DanhSachHaoPhi = congTacChuan.DanhSachHaoPhi;

                decimal dgVL = 0;
                decimal dgNC = 0;
                decimal dgMay = 0;

                // Tách riêng vật liệu khác (tỷ lệ %) để tính sau cùng
                var cacHaoPhiBinhThuong = congTacChuan.DanhSachHaoPhi
                    .Where(hp => !(hp.LoaiHaoPhi == LoaiHaoPhi.VL && (hp.DonVi == "%" || hp.TenHaoPhi.ToLower().Contains("vật liệu khác"))));
                var cacHaoPhiKhac = congTacChuan.DanhSachHaoPhi
                    .Where(hp => hp.LoaiHaoPhi == LoaiHaoPhi.VL && (hp.DonVi == "%" || hp.TenHaoPhi.ToLower().Contains("vật liệu khác")));

                foreach (var hp in cacHaoPhiBinhThuong)
                {
                    decimal giaHT = 0;
                    if (hp.LoaiHaoPhi == LoaiHaoPhi.VL && dictVL.TryGetValue(hp.MaHieuHP, out var gVL))
                    {
                        giaHT = gVL;
                    }
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.NC && dictNC.TryGetValue(hp.MaHieuHP, out var gNC))
                    {
                        giaHT = gNC;
                    }
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.MAY && dictMay.TryGetValue(hp.MaHieuHP, out var gMay))
                    {
                        giaHT = gMay;
                    }
                    
                    // Tính đơn giá = Hao phí (Định mức x Hệ số) x Giá Hiện Trường
                    decimal chiPhi = hp.DinhMuc * hp.HeSo * giaHT;
                    
                    if (hp.LoaiHaoPhi == LoaiHaoPhi.VL) dgVL += chiPhi;
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.NC) dgNC += chiPhi;
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.MAY) dgMay += chiPhi;
                }

                // Tính vật liệu khác (theo %)
                foreach (var hp in cacHaoPhiKhac)
                {
                    // Vật liệu khác = (Tổng chi phí vật liệu chính) * Định mức %
                    decimal chiPhiVlKhac = dgVL * (hp.DinhMuc / 100m);
                    dgVL += chiPhiVlKhac;
                }

                ct.DonGiaVL = dgVL;
                ct.DonGiaNC = dgNC;
                ct.DonGiaMay = dgMay;
            }
        }
    }
}
