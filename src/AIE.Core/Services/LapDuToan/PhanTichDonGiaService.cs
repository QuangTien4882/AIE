using AIE.Core.Models;

namespace AIE.Core.Services.LapDuToan;

/// <summary>
/// Service phân tích đơn giá chi tiết cho 1 công tác xây dựng.
/// Tính đơn giá = Σ(hao phí × đơn giá) cho VL, NC, Máy.
/// </summary>
public class PhanTichDonGiaService
{
    /// <summary>
    /// Phân tích đơn giá chi tiết cho 1 công tác.
    /// </summary>
    /// <param name="congTac">Công tác XD với danh sách hao phí</param>
    /// <param name="giaVatLieu">Bảng giá VL (key = MaVL)</param>
    /// <param name="giaNhanCong">Bảng giá NC (key = MaNC)</param>
    /// <param name="giaMay">Bảng giá máy (key = MaMay)</param>
    /// <returns>Kết quả PTĐG</returns>
    public KetQuaPhanTichDonGia PhanTich(
        CongTacXayDung congTac,
        IDictionary<string, decimal> giaVatLieu,
        IDictionary<string, decimal> giaNhanCong,
        IDictionary<string, decimal> giaMay)
    {
        var ketQua = new KetQuaPhanTichDonGia
        {
            MaHieu = congTac.MaHieu,
            TenCongTac = congTac.TenCongTac,
            DonVi = congTac.DonVi
        };

        foreach (var hp in congTac.DanhSachHaoPhi)
        {
            var chiTiet = new ChiTietHaoPhi
            {
                MaHieuHP = hp.MaHieuHP,
                TenHaoPhi = hp.TenHaoPhi,
                DonVi = hp.DonVi,
                DinhMuc = hp.DinhMuc,
                HeSo = hp.HeSo
            };

            switch (hp.LoaiHaoPhi)
            {
                case Enums.LoaiHaoPhi.VL:
                    chiTiet.DonGia = giaVatLieu.TryGetValue(hp.MaHieuHP, out var giaVL) ? giaVL : 0;
                    ketQua.ChiTietVatLieu.Add(chiTiet);
                    break;

                case Enums.LoaiHaoPhi.NC:
                    chiTiet.DonGia = giaNhanCong.TryGetValue(hp.MaHieuHP, out var giaNC) ? giaNC : 0;
                    ketQua.ChiTietNhanCong.Add(chiTiet);
                    break;

                case Enums.LoaiHaoPhi.MAY:
                    chiTiet.DonGia = giaMay.TryGetValue(hp.MaHieuHP, out var giaM) ? giaM : 0;
                    ketQua.ChiTietMay.Add(chiTiet);
                    break;
            }
        }

        return ketQua;
    }

    /// <summary>
    /// Phân tích đơn giá cho nhiều công tác cùng lúc.
    /// </summary>
    public List<KetQuaPhanTichDonGia> PhanTichNhieu(
        IEnumerable<CongTacXayDung> dsCongTac,
        IDictionary<string, decimal> giaVatLieu,
        IDictionary<string, decimal> giaNhanCong,
        IDictionary<string, decimal> giaMay)
    {
        return dsCongTac
            .Select(ct => PhanTich(ct, giaVatLieu, giaNhanCong, giaMay))
            .ToList();
    }
}
