using AIE.Core.Models;
using AIE.Core.Enums;
using AIE.Data.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace AIE.ExcelAddIn.Services;

public class ThamDinhEngine
{
    private readonly CongTacRepository _repository;
    private readonly VatLieuRepository _vlRepo;
    private readonly NhanCongRepository _ncRepo;
    private readonly MayThiCongRepository _mayRepo;

    public ThamDinhEngine(
        CongTacRepository repository,
        VatLieuRepository vlRepo,
        NhanCongRepository ncRepo,
        MayThiCongRepository mayRepo)
    {
        _repository = repository;
        _vlRepo = vlRepo;
        _ncRepo = ncRepo;
        _mayRepo = mayRepo;
    }

    public List<KetQuaCongTacThamDinh> KiemTra(List<CongTacThamDinh> duToanList)
    {
        var ketQua = new List<KetQuaCongTacThamDinh>();

        foreach (var ctDuToan in duToanList)
        {
            var kq = new KetQuaCongTacThamDinh { DuToan = ctDuToan };
            var ctChuan = _repository.GetByMaHieu(ctDuToan.MaHieu);

            if (ctChuan == null)
            {
                kq.DanhSachSaiLech.Add(new SaiLechDinhMuc
                {
                    LoaiLoi = "Mã không tồn tại",
                    MoTa = $"Mã hiệu '{ctDuToan.MaHieu}' không tồn tại trong CSDL định mức.",
                    SoDongExcel = ctDuToan.SoDongExcel
                });
                ketQua.Add(kq);
                continue;
            }

            kq.DinhMucChuan = ctChuan;

            // Tạo bản sao danh sách chuẩn để đánh dấu những mục đã map
            var dsHpChuan = ctChuan.DanhSachHaoPhi.ToList();

            foreach (var hpDuToan in ctDuToan.DanhSachHaoPhi)
            {
                var hpChuanMatched = TimHaoPhiTuongDuong(hpDuToan, dsHpChuan);
                
                var saiLech = new SaiLechDinhMuc
                {
                    HaoPhiDuToan = hpDuToan,
                    LoaiHP = hpDuToan.Loai,
                    SoDongExcel = hpDuToan.SoDongExcel
                };

                if (hpChuanMatched == null)
                {
                    saiLech.LoaiLoi = "Hao phí thừa / Không khớp";
                    saiLech.MoTa = $"Dự toán có '{hpDuToan.TenHaoPhi}' nhưng TT38 không có (hoặc tên không khớp).";
                }
                else
                {
                    dsHpChuan.Remove(hpChuanMatched); // Đã map
                    
                    saiLech.HaoPhiChuan = hpChuanMatched;
                    saiLech.DonGiaChuan = GetDonGiaChuan(hpChuanMatched);

                    // Kiểm tra định mức
                    if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                    {
                        saiLech.LoaiLoi = "Sai định mức";
                        saiLech.MoTa = $"'{hpDuToan.TenHaoPhi}': DT = {hpDuToan.DinhMuc:G}, TT38 = {hpChuanMatched.DinhMuc:G}";
                    }

                    // Kiểm tra đơn giá
                    if (saiLech.DonGiaChuan.HasValue && Math.Abs(saiLech.ChenhLechDonGia) > 1)
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai đơn giá";
                        else saiLech.LoaiLoi += ", Sai đơn giá";
                    }

                    // Kiểm tra đơn vị
                    if (!string.Equals(hpDuToan.DonVi, hpChuanMatched.DonVi, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai đơn vị";
                        else saiLech.LoaiLoi += ", Sai đơn vị";
                    }
                }
                
                kq.DanhSachSaiLech.Add(saiLech);
            }

            // Những hao phí chuẩn còn lại (chưa được map) là bị thiếu
            foreach (var hpThieu in dsHpChuan)
            {
                kq.DanhSachSaiLech.Add(new SaiLechDinhMuc
                {
                    LoaiLoi = "Thiếu hao phí",
                    MoTa = $"TT38 có '{hpThieu.TenHaoPhi}' nhưng Dự toán bị thiếu.",
                    LoaiHP = hpThieu.LoaiHaoPhi,
                    SoDongExcel = ctDuToan.SoDongExcel, // Báo ở dòng công tác vì không có dòng hao phí
                    HaoPhiChuan = hpThieu,
                    DonGiaChuan = GetDonGiaChuan(hpThieu)
                });
            }

            ketQua.Add(kq);
        }

        return ketQua;
    }
    
    private decimal? GetDonGiaChuan(HaoPhi hpChuan)
    {
        if (string.IsNullOrEmpty(hpChuan.MaHieuHP)) return null;

        if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.VL)
        {
            var vl = _vlRepo.GetByMa(hpChuan.MaHieuHP);
            return vl?.DonGia;
        }
        else if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.NC)
        {
            var nc = _ncRepo.GetByMa(hpChuan.MaHieuHP);
            return nc?.DonGia;
        }
        else if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.MAY)
        {
            var may = _mayRepo.GetByMa(hpChuan.MaHieuHP);
            return may?.DonGia;
        }
        return null;
    }

    private HaoPhi? TimHaoPhiTuongDuong(HaoPhiThamDinh hpDuToan, List<HaoPhi> dsHpChuan)
    {
        string tenDt = ChuanHoaTen(hpDuToan.TenHaoPhi);

        // Ưu tiên tìm khớp chính xác hoàn toàn
        var exactMatch = dsHpChuan.FirstOrDefault(x => ChuanHoaTen(x.TenHaoPhi) == tenDt && x.LoaiHaoPhi == hpDuToan.Loai);
        if (exactMatch != null) return exactMatch;

        // Nếu không có, tìm kiếm gần đúng (chứa chuỗi)
        var containsMatch = dsHpChuan.FirstOrDefault(x =>
            (ChuanHoaTen(x.TenHaoPhi).Contains(tenDt) || tenDt.Contains(ChuanHoaTen(x.TenHaoPhi)))
            && x.LoaiHaoPhi == hpDuToan.Loai);

        return containsMatch;
    }

    private string ChuanHoaTen(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var result = input.Trim().ToLower();
        if (result.StartsWith("-")) result = result.Substring(1).Trim();
        // Loại bỏ khoảng trắng thừa
        result = Regex.Replace(result, @"\s+", " ");
        return result;
    }
}
