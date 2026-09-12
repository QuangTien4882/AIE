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

                // Tách riêng các hao phí tỷ lệ % (vật liệu khác, máy khác, nhân công khác) để tính sau cùng
                bool IsPercentageHaoPhi(HaoPhi hp) =>
                    hp.DonVi == "%" || 
                    (hp.TenHaoPhi != null && hp.TenHaoPhi.ToLower().Contains("khác"));

                var cacHaoPhiBinhThuong = congTacChuan.DanhSachHaoPhi.Where(hp => !IsPercentageHaoPhi(hp));
                var cacHaoPhiKhac = congTacChuan.DanhSachHaoPhi.Where(hp => IsPercentageHaoPhi(hp));

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

                // Tính các hao phí khác (theo % của tổng chi phí tương ứng)
                foreach (var hp in cacHaoPhiKhac)
                {
                    decimal rate = (hp.DinhMuc * hp.HeSo) / 100m;
                    if (hp.LoaiHaoPhi == LoaiHaoPhi.VL)
                    {
                        dgVL += dgVL * rate;
                    }
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.MAY)
                    {
                        dgMay += dgMay * rate;
                    }
                    else if (hp.LoaiHaoPhi == LoaiHaoPhi.NC)
                    {
                        dgNC += dgNC * rate;
                    }
                }

                ct.DonGiaVL = dgVL;
                ct.DonGiaNC = dgNC;
                ct.DonGiaMay = dgMay;
            }
        }
    }
}
