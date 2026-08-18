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

    public List<KetQuaCongTacThamDinh> KiemTra(List<CongTacThamDinh> duToanList, int? boDonGiaId = null)
    {
        var ketQua = new List<KetQuaCongTacThamDinh>();
        
        BoDonGiaRepository? bdgRepo = null;
        if (boDonGiaId.HasValue)
        {
            var db = new AIE.Data.DatabaseManager();
            bdgRepo = new BoDonGiaRepository(db.Context);
        }

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
                var hpChuanMatched = TimHaoPhiTuongDuong(hpDuToan, dsHpChuan, ctChuan);
                
                var saiLech = new SaiLechDinhMuc
                {
                    HaoPhiDuToan = hpDuToan,
                    LoaiHP = hpDuToan.Loai,
                    SoDongExcel = hpDuToan.SoDongExcel
                };

                if (hpChuanMatched == null)
                {
                    saiLech.LoaiLoi = "Hao phí thừa";
                    saiLech.MoTa = $"Dự toán có '{hpDuToan.TenHaoPhi}' nhưng TT38 không có (hoặc tên không khớp).";
                }
                else
                {
                    dsHpChuan.Remove(hpChuanMatched); // Đã map
                    
                    saiLech.HaoPhiChuan = hpChuanMatched;
                    saiLech.DonGiaChuan = GetDonGiaChuan(hpChuanMatched, boDonGiaId, bdgRepo);

                    // Kiểm tra khác tên gọi (do Smart Mapping)
                    string tenDt = ChuanHoaTen(hpDuToan.TenHaoPhi);
                    string tenCh = ChuanHoaTen(hpChuanMatched.TenHaoPhi);
                    if (tenDt != tenCh && !tenDt.Contains(tenCh) && !tenCh.Contains(tenDt))
                    {
                        saiLech.LoaiLoi = "Khác tên gọi";
                        saiLech.MoTa = $"DT: '{hpDuToan.TenHaoPhi}' | TT38: '{hpChuanMatched.TenHaoPhi}'";
                    }

                    // Kiểm tra định mức
                    if (Math.Abs(saiLech.ChenhLechDinhMuc) > 0.0001m)
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai định mức";
                        else saiLech.LoaiLoi += ", Sai định mức";
                        
                        if (string.IsNullOrEmpty(saiLech.MoTa)) saiLech.MoTa = $"DT = {hpDuToan.DinhMuc:G}, TT38 = {hpChuanMatched.DinhMuc:G}";
                        else saiLech.MoTa += $"\nSai ĐM: DT = {hpDuToan.DinhMuc:G}, TT38 = {hpChuanMatched.DinhMuc:G}";
                    }

                    // Bỏ Kiểm tra đơn giá ở phần định mức này vì sẽ tách riêng
                    
                    // Kiểm tra Đơn vị
                    if (!string.Equals(hpDuToan.DonVi, hpChuanMatched.DonVi, StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(saiLech.LoaiLoi)) saiLech.LoaiLoi = "Sai Đơn vị";
                        else saiLech.LoaiLoi += ", Sai Đơn vị";
                    }
                }
                
                kq.DanhSachSaiLech.Add(saiLech);
            }

            // Những hao phí chuẩn còn lại (chưa được map) là bị thiếu
            foreach (var hpThieu in dsHpChuan)
            {
                var sl = new SaiLechDinhMuc
                {
                    LoaiLoi = "Thiếu hao phí",
                    MoTa = $"TT38 có '{hpThieu.TenHaoPhi}' nhưng Dự toán bị thiếu.",
                    LoaiHP = hpThieu.LoaiHaoPhi,
                    SoDongExcel = ctDuToan.SoDongExcel, // Báo ở dòng công tác vì không có dòng hao phí
                    HaoPhiChuan = hpThieu
                };
                
                var giaChuan = GetDonGiaChuan(hpThieu, boDonGiaId, bdgRepo);
                sl.DonGiaChuan = giaChuan;

                kq.DanhSachSaiLech.Add(sl);
            }

            ketQua.Add(kq);
        }

        return ketQua;
    }

    public List<VatTuGiaModel> TrichXuatVatTu(List<KetQuaCongTacThamDinh> ketQuaDinhMuc)
    {
        var dict = new Dictionary<string, VatTuGiaModel>();

        foreach (var kq in ketQuaDinhMuc)
        {
            foreach (var saiLech in kq.DanhSachSaiLech)
            {
                if (saiLech.HaoPhiDuToan == null) continue;
                
                // Dùng tên chuẩn hóa làm key để gộp những vật tư giống tên
                string key = ChuanHoaTen(saiLech.HaoPhiDuToan.TenHaoPhi) + "_" + saiLech.HaoPhiDuToan.Loai.ToString();

                if (!dict.ContainsKey(key))
                {
                    dict[key] = new VatTuGiaModel
                    {
                        TenVatTu = saiLech.HaoPhiDuToan.TenHaoPhi,
                        DonVi = saiLech.HaoPhiDuToan.DonVi,
                        LoaiHP = saiLech.HaoPhiDuToan.Loai,
                        GiaDuToan = saiLech.HaoPhiDuToan.DonGia,
                        MaHieu = saiLech.HaoPhiChuan?.MaHieuHP ?? string.Empty,
                        GiaChuan = saiLech.DonGiaChuan
                    };
                }
            }
        }

        return dict.Values.OrderBy(x => x.LoaiHP).ThenBy(x => x.TenVatTu).ToList();
    }
    
    private decimal? GetDonGiaChuan(HaoPhi hpChuan, int? boDonGiaId = null, BoDonGiaRepository? bdgRepo = null)
    {
        if (string.IsNullOrEmpty(hpChuan.MaHieuHP)) return null;

        if (boDonGiaId.HasValue && bdgRepo != null)
        {
            if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.VL)
            {
                var vl = bdgRepo.GetGiaVL(boDonGiaId.Value).FirstOrDefault(x => x.MaVL == hpChuan.MaHieuHP);
                if (vl != null) return vl.GiaHienTruong;
            }
            else if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.NC)
            {
                var nc = bdgRepo.GetGiaNC(boDonGiaId.Value).FirstOrDefault(x => x.MaNC == hpChuan.MaHieuHP);
                if (nc != null) return nc.DonGia;
            }
            else if (hpChuan.LoaiHaoPhi == LoaiHaoPhi.MAY)
            {
                var may = bdgRepo.GetGiaMay(boDonGiaId.Value).FirstOrDefault(x => x.MaMay == hpChuan.MaHieuHP);
                if (may != null) return may.DonGia;
            }
        }

        // Fallback to global catalog
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

    private HaoPhi? TimHaoPhiTuongDuong(HaoPhiThamDinh hpDuToan, List<HaoPhi> dsHpChuan, CongTacXayDung ctChuan)
    {
        string tenDt = ChuanHoaTen(hpDuToan.TenHaoPhi);

        // 1. Khớp chính xác hoàn toàn
        var exactMatch = dsHpChuan.FirstOrDefault(x => ChuanHoaTen(x.TenHaoPhi) == tenDt && x.LoaiHaoPhi == hpDuToan.Loai);
        if (exactMatch != null) return exactMatch;

        // 2. Khớp gần đúng (chứa chuỗi)
        var containsMatch = dsHpChuan.FirstOrDefault(x =>
            (ChuanHoaTen(x.TenHaoPhi).Contains(tenDt) || tenDt.Contains(ChuanHoaTen(x.TenHaoPhi)))
            && x.LoaiHaoPhi == hpDuToan.Loai);
            
        if (containsMatch != null) return containsMatch;

        // 3. Khớp theo mức độ trùng lặp từ vựng (Word Overlap)
        var wordsDt = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
        HaoPhi? bestMatch = null;
        int maxOverlap = 0;

        foreach (var hp in dsHpChuan.Where(x => x.LoaiHaoPhi == hpDuToan.Loai))
        {
            var wordsCh = ChuanHoaTen(hp.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            int overlap = wordsDt.Intersect(wordsCh).Count();
            
            if (overlap > maxOverlap)
            {
                maxOverlap = overlap;
                bestMatch = hp;
            }
        }

        // Bắt cặp nếu trùng ít nhất 2 từ (VD: "Máy đào", "Máy đầm", "Vật liệu")
        // Nếu tên quá ngắn (chỉ 1 từ) thì cần trùng 1 từ
        int minRequiredOverlap = wordsDt.Length <= 1 ? 1 : 2;
        if (bestMatch != null && maxOverlap >= minRequiredOverlap)
        {
            return bestMatch;
        }

        // 4. Smart Mapping: Nếu công tác chỉ có 1 hao phí loại này, và chuẩn cũng có 1 -> Bắt cặp
        // Nhưng BẮT BUỘC phải có ít nhất 1 từ trùng để tránh ghép nhầm (VD: "Nhựa bitum" với "Bột đá")
        var countChuanLoaiNay = ctChuan.DanhSachHaoPhi.Count(x => x.LoaiHaoPhi == hpDuToan.Loai);
        if (countChuanLoaiNay == 1)
        {
            var smartMatch = dsHpChuan.FirstOrDefault(x => x.LoaiHaoPhi == hpDuToan.Loai);
            if (smartMatch != null)
            {
                // Kiểm tra có ít nhất 1 từ trùng
                var w1 = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                var w2 = ChuanHoaTen(smartMatch.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (w1.Intersect(w2).Any()) return smartMatch;
            }
        }
        
        // 5. Fallback cuối cùng: Nếu chỉ còn 1 hao phí TT38 chưa map cùng loại, bắt cặp NẾU có từ trùng.
        // Tuyệt đối KHÔNG bắt cặp khi tên hoàn toàn khác nhau → phải để thành "Hao phí thừa".
        var remainingChuanLoaiNay = dsHpChuan.Where(x => x.LoaiHaoPhi == hpDuToan.Loai).ToList();
        if (remainingChuanLoaiNay.Count == 1)
        {
            var candidate = remainingChuanLoaiNay.First();
            var w1 = tenDt.Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            var w2 = ChuanHoaTen(candidate.TenHaoPhi).Split(new[] { ' ', '-', ',', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            if (w1.Intersect(w2).Any()) return candidate;
        }

        return null;
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
