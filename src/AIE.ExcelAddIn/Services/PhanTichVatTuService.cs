using AIE.Core.Models;
using AIE.Core.Enums;
using AIE.Data.Repositories;
using System.Collections.Generic;
using System.Linq;

namespace AIE.ExcelAddIn.Services
{
    public class PhanTichVatTuService
    {
        private readonly CongTacRepository _congTacRepo;
        private readonly MayThiCongRepository _mayRepo;
        private readonly DinhMucCaMayRepository _dinhMucMayRepo;
        private readonly VatLieuRepository _vlRepo;
        private readonly NhanCongRepository _ncRepo;
        
        public PhanTichVatTuService(
            CongTacRepository congTacRepo, 
            MayThiCongRepository mayRepo,
            DinhMucCaMayRepository dinhMucMayRepo,
            VatLieuRepository vlRepo, 
            NhanCongRepository ncRepo)
        {
            _congTacRepo = congTacRepo;
            _mayRepo = mayRepo;
            _dinhMucMayRepo = dinhMucMayRepo;
            _vlRepo = vlRepo;
            _ncRepo = ncRepo;
        }

        public void PhanTich(DuToan duToan)
        {
            var bangTongHop = new BangTongHopVatTu();
            
            // Load fuel prices
            if (duToan.BoDonGiaId.HasValue)
            {
                var db = new AIE.Data.DatabaseManager();
                var boRepo = new BoDonGiaRepository(db.Context);
                var bo = boRepo.GetById(duToan.BoDonGiaId.Value);
                if (bo != null)
                {
                    bangTongHop.GiaXang = bo.GiaXang;
                    bangTongHop.GiaDiezel = bo.GiaDiezel;
                    bangTongHop.GiaDien = bo.GiaDien;
                }
            }
            else
            {
                // Fallback to JSON if no BoDonGia
                try
                {
                    var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                    var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                    if (System.IO.File.Exists(fuelFile))
                    {
                        var json = System.IO.File.ReadAllText(fuelFile);
                        dynamic doc = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                        bangTongHop.GiaXang = (decimal)doc.GiaXang;
                        bangTongHop.GiaDiezel = (decimal)doc.GiaDiezel;
                        bangTongHop.GiaDien = (decimal)doc.GiaDien;
                    }
                }
                catch { }
            }

            var dictVL = new Dictionary<string, VatLieuHienTruong>();
            var dictNC = new Dictionary<string, NhanCongHienTruong>();
            var dictMay = new Dictionary<string, MayThiCongHienTruong>();

            foreach (var hm in duToan.DanhSachHangMuc)
            {
                foreach (var ct in hm.DanhSachCongTac)
                {
                    // Lấy định mức chuẩn từ DB
                    var congTacChuan = _congTacRepo.GetByMaHieu(ct.MaHieu);
                    if (congTacChuan == null) continue;

                    foreach (var hp in congTacChuan.DanhSachHaoPhi)
                    {
                        var tongKhoiLuong = ct.KhoiLuong * hp.DinhMuc * hp.HeSo;
                        
                        if (hp.LoaiHaoPhi == LoaiHaoPhi.VL)
                        {
                            if (hp.DonVi == "%" || hp.TenHaoPhi.ToLower().Contains("vật liệu khác"))
                            {
                                continue;
                            }

                            if (!dictVL.ContainsKey(hp.MaHieuHP))
                            {
                                var vlMaster = _vlRepo.GetByMa(hp.MaHieuHP);
                                dictVL[hp.MaHieuHP] = new VatLieuHienTruong
                                {
                                    MaVatTu = hp.MaHieuHP,
                                    TenVatTu = hp.TenHaoPhi,
                                    DonVi = hp.DonVi,
                                    GiaGoc = vlMaster?.DonGia ?? 0,
                                    CuocVanChuyen = vlMaster?.CuocVanChuyen ?? 0
                                };
                            }
                            dictVL[hp.MaHieuHP].TongKhoiLuong += tongKhoiLuong;
                        }
                        else if (hp.LoaiHaoPhi == LoaiHaoPhi.NC)
                        {
                            if (!dictNC.ContainsKey(hp.MaHieuHP))
                            {
                                var ncMaster = _ncRepo.GetByMa(hp.MaHieuHP);
                                dictNC[hp.MaHieuHP] = new NhanCongHienTruong
                                {
                                    MaVatTu = hp.MaHieuHP,
                                    TenVatTu = hp.TenHaoPhi,
                                    DonVi = hp.DonVi,
                                    GiaGoc = ncMaster?.DonGia ?? 0,
                                    GiaHienTruong = ncMaster?.DonGia ?? 0 // Mặc định bằng giá gốc
                                };
                            }
                            dictNC[hp.MaHieuHP].TongKhoiLuong += tongKhoiLuong;
                        }
                        else if (hp.LoaiHaoPhi == LoaiHaoPhi.MAY)
                        {
                            if (hp.TenHaoPhi.ToLower().Contains("máy khác") || hp.MaHieuHP == "M7016")
                            {
                                continue;
                            }

                            if (!dictMay.ContainsKey(hp.MaHieuHP))
                            {
                                var mayMaster = _mayRepo.GetByMa(hp.MaHieuHP);
                                var dmTT37 = _dinhMucMayRepo.GetByMaMay(hp.MaHieuHP);
                                
                                dictMay[hp.MaHieuHP] = new MayThiCongHienTruong
                                {
                                    MaVatTu = hp.MaHieuHP,
                                    TenVatTu = hp.TenHaoPhi,
                                    DonVi = hp.DonVi,
                                    DinhMuc = dmTT37,
                                    GiaGoc = mayMaster?.DonGia ?? 0
                                };
                            }
                            dictMay[hp.MaHieuHP].TongKhoiLuong += tongKhoiLuong;
                        }
                    }
                }
            }

            bangTongHop.DanhSachVatLieu = dictVL.Values.OrderBy(x => x.MaVatTu).ToList();
            bangTongHop.DanhSachNhanCong = dictNC.Values.OrderBy(x => x.MaVatTu).ToList();
            bangTongHop.DanhSachMay = dictMay.Values.OrderBy(x => x.MaVatTu).ToList();
            
            // Hàm tính lại giá máy dựa trên thông số nhiên liệu và định mức
            TinhGiaMayThiCong(bangTongHop);

            duToan.BangTongHop = bangTongHop;
        }

            public void TinhGiaMayThiCong(BangTongHopVatTu bangTongHop)
        {
            var soCaNam = 250m; // Số ca năm (theo chuẩn thường là 250 ca/năm)

            // Hệ số nhiên liệu phụ theo TT37
            decimal hsXang = 1.02m;
            decimal hsDiezel = 1.03m;
            decimal hsDien = 1.05m;

            // Materialize danh sách nhân công 1 lần duy nhất (tránh lỗi Dapper deferred execution)
            var danhSachNhanCong = _ncRepo.GetAll().ToList();
            
            foreach (var may in bangTongHop.DanhSachMay)
            {
                if (may.DinhMuc == null)
                {
                    // Nếu không có định mức TT37, giữ nguyên giá gốc
                    may.GiaHienTruong = may.GiaGoc;
                    continue;
                }

                var dm = may.DinhMuc;
                may.NguyenGia = dm.NguyenGia;
                
                // Giá trị thu hồi theo TT37: Nguyên giá >= 30 triệu thì G_TH = 10% Nguyên giá
                decimal g_th = dm.NguyenGia >= 30000000m ? dm.NguyenGia * 0.1m : 0m;
                
                // Khấu hao = (Nguyên giá - G_TH) x Tỷ lệ khấu hao / Số ca năm
                may.ChiPhiKhauHao = ((dm.NguyenGia - g_th) * (dm.KhauHao / 100m)) / soCaNam;
                
                // Sửa chữa = Nguyên giá x Tỷ lệ SC / Số ca năm
                may.ChiPhiSuaChua = dm.NguyenGia * (dm.SuaChua / 100m) / soCaNam;
                
                // Chi phí khác = Nguyên giá x Tỷ lệ Khác / Số ca năm
                may.ChiPhiKhac = dm.NguyenGia * (dm.ChiPhiKhac / 100m) / soCaNam;

                // Nhiên liệu = (Lượng xăng * Giá Xăng * HS) + (Lượng Diezel * Giá Diezel * HS) + (Lượng điện * Giá Điện * HS)
                may.ChiPhiNhiemLieu = dm.DinhMucXang * bangTongHop.GiaXang * hsXang
                                    + dm.DinhMucDiezel * bangTongHop.GiaDiezel * hsDiezel
                                    + dm.DinhMucDien * bangTongHop.GiaDien * hsDien;

                // Chi phí thợ lái máy tự động lấy từ DB theo NhomNhanCong của máy
                var nc = danhSachNhanCong.FirstOrDefault(x => x.Nhom == dm.NhomNhanCong);
                var giaNhanCong = nc?.DonGia ?? 0;
                may.ChiPhiNhanCong = dm.SoLuongNhanCong * giaNhanCong;
            }
        }

        public void SaveGia(BangTongHopVatTu bangTongHop)
        {
            // Lưu Vật Liệu (Cập nhật Giá gốc và Cước vận chuyển)
            // Trong DB, VatLieu có DonGia. Tạm thời coi DonGia là Giá Gốc. 
            // Cước VC có thể cần lưu thêm nếu DB có cột này.
            // Giả sử bảng VatLieu chỉ có DonGia.
            foreach (var vl in bangTongHop.DanhSachVatLieu)
            {
                var dbVl = _vlRepo.GetByMa(vl.MaVatTu);
                if (dbVl != null)
                {
                    dbVl.DonGia = vl.GiaGoc;
                    dbVl.CuocVanChuyen = vl.CuocVanChuyen;
                    _vlRepo.Upsert(dbVl);
                }
            }

            // Save fuel prices to JSON
            try
            {
                var settingsDir = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
                if (!System.IO.Directory.Exists(settingsDir)) System.IO.Directory.CreateDirectory(settingsDir);
                var fuelFile = System.IO.Path.Combine(settingsDir, "fuel_prices.json");
                var fuelData = new { bangTongHop.GiaXang, bangTongHop.GiaDiezel, bangTongHop.GiaDien };
                System.IO.File.WriteAllText(fuelFile, Newtonsoft.Json.JsonConvert.SerializeObject(fuelData));
            }
            catch { }

            // Lưu Nhân Công
            foreach (var nc in bangTongHop.DanhSachNhanCong)
            {
                var dbNc = _ncRepo.GetByMa(nc.MaVatTu);
                if (dbNc != null)
                {
                    dbNc.DonGia = nc.GiaGoc; // Giá gốc = giá nhân công nhập vào
                    _ncRepo.Upsert(dbNc);
                }
            }

            // Lưu Máy
            // Vì máy có phần Giá Xăng/Diezel/Điện không được lưu riêng cho từng máy, 
            // Giá gốc của máy được nhập thủ công? 
            // Hoặc giá máy được tính tự động từ nhiên liệu?
            // Thực tế Giá gốc có thể update lại cho MayThiCong
            foreach (var m in bangTongHop.DanhSachMay)
            {
                var dbMay = _mayRepo.GetByMa(m.MaVatTu);
                if (dbMay != null)
                {
                    dbMay.DonGia = m.GiaGoc;
                    _mayRepo.Upsert(dbMay);
                }
            }
        }
    }
}
