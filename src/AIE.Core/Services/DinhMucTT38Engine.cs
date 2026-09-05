using System;
using System.Collections.Generic;
using System.Linq;
using AIE.Core.Models;

namespace AIE.Core.Services
{
    /// <summary>
    /// Engine tra cứu, nội suy tỷ lệ định mức và tính toán Tổng mức đầu tư (Bảng 1.2 TT 36)
    /// & Tổng hợp dự toán công trình (Bảng 2.1 TT 36) theo Thông tư 38/2026 và các văn bản BTC.
    /// </summary>
    public static class DinhMucTT38Engine
    {
        /// <summary>
        /// Tạo mới một BangTongHopKinhPhiModel với đầy đủ các danh mục chi phí chuẩn
        /// </summary>
        public static BangTongHopKinhPhiModel TaoBangKinhPhiMacDinh(
            string loaiCT = "Dân dụng",
            string capCT = "Cấp III",
            int soBuocTK = 2,
            decimal chiPhiXD = 10000000000m,   // 10 tỷ
            decimal chiPhiTB = 0m,
            decimal chiPhiBT = 0m,
            decimal chiPhiNhaTam = 0m)
        {
            var model = new BangTongHopKinhPhiModel
            {
                LoaiCongTrinh = loaiCT,
                CapCongTrinh = capCT,
                SoBuocThietKe = soBuocTK,
                ChiPhiXDTruocThue = chiPhiXD,
                ChiPhiNhaTamTruocThue = chiPhiNhaTam,
                ChiPhiTBTruocThue = chiPhiTB,
                ChiPhiBTTruocThue = chiPhiBT
            };

            model.Items = KhoiTaoDanhSachKhoanMucChuan();
            if (chiPhiTB > 0)
            {
                var itemGSTB = model.Items.FirstOrDefault(x => x.MaChiPhi == "TV_GS_TB");
                if (itemGSTB != null) itemGSTB.IsActive = true;
                var itemTB = model.Items.FirstOrDefault(x => x.Nhom == NhomChiPhi.ChiPhiThietBi);
                if (itemTB != null) itemTB.IsActive = true;
            }
            CapNhatToanBoDinhMucVaTinhToan(model);
            return model;
        }

        /// <summary>
        /// Khởi tạo khung các khoản mục chi phí chuẩn theo Bảng 1.2 và Bảng 2.1
        /// </summary>
        public static List<ChiPhiKinhPhiItem> KhoiTaoDanhSachKhoanMucChuan()
        {
            var list = new List<ChiPhiKinhPhiItem>
            {
                // I. Chi phí bồi thường, hỗ trợ và tái định cư
                new ChiPhiKinhPhiItem
                {
                    STT = "I",
                    MaChiPhi = "G_BT",
                    TenChiPhi = "Chi phí bồi thường, hỗ trợ và tái định cư",
                    Nhom = NhomChiPhi.BoiThuong_TDC,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiBoiThuong,
                    KyHieu = "G_BT,TĐC",
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Phương án bồi thường GPMB được duyệt",
                    IsActive = false, // Mặc định tắt nếu không có GPMB
                    IsReadOnly = true
                },

                // II. Chi phí xây dựng (Tách thành 2 mục theo quy định: Chi phí xây dựng & Chi phí nhà tạm)
                new ChiPhiKinhPhiItem
                {
                    STT = "2.1",
                    MaChiPhi = "G_XD",
                    TenChiPhi = "Chi phí xây dựng",
                    Nhom = NhomChiPhi.ChiPhiXayDung,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gxd",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng tổng hợp chi phí xây dựng (Dòng IV Bảng 3.8)",
                    IsActive = true,
                    IsReadOnly = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "2.2",
                    MaChiPhi = "G_NHA_TAM",
                    TenChiPhi = "Chi phí nhà tạm để ở và điều hành thi công",
                    Nhom = NhomChiPhi.ChiPhiXayDung,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gnt",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng tổng hợp chi phí xây dựng (Dòng VII Bảng 3.8)",
                    IsActive = true,
                    IsReadOnly = true
                },

                // III. Chi phí thiết bị
                new ChiPhiKinhPhiItem
                {
                    STT = "III",
                    MaChiPhi = "G_TB",
                    TenChiPhi = "Chi phí thiết bị",
                    Nhom = NhomChiPhi.ChiPhiThietBi,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    KyHieu = "Gtb",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán mua sắm thiết bị (Bảng 2.2)",
                    IsActive = true,
                    IsReadOnly = true
                },

                // IV. Chi phí quản lý dự án
                new ChiPhiKinhPhiItem
                {
                    STT = "IV",
                    MaChiPhi = "G_QLDA",
                    TenChiPhi = "Chi phí quản lý dự án",
                    Nhom = NhomChiPhi.QuanLyDuAn,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gqlda",
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Bảng 1.1 Thông tư 38/2026/TT-BXD",
                    IsActive = true,
                    IsReadOnly = true
                },

                // V. Chi phí tư vấn đầu tư xây dựng
                new ChiPhiKinhPhiItem
                {
                    STT = "5.1",
                    MaChiPhi = "TV_KS",
                    TenChiPhi = "Chi phí khảo sát xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv1",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán chi phí khảo sát riêng",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.2",
                    MaChiPhi = "TV_FS_KTKT",
                    TenChiPhi = "Chi phí lập Báo cáo nghiên cứu khả thi (FS) / Báo cáo KT-KT",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gtv2",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.2 / Bảng 2.4 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.3",
                    MaChiPhi = "TV_TK",
                    TenChiPhi = "Chi phí thiết kế xây dựng công trình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv3",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.7 - 2.16 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.4",
                    MaChiPhi = "TV_TT_TK",
                    TenChiPhi = "Chi phí thẩm tra thiết kế xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv4",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.19 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.5",
                    MaChiPhi = "TV_TT_DT",
                    TenChiPhi = "Chi phí thẩm tra dự toán xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv5",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.20 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.6",
                    MaChiPhi = "TV_HSMT",
                    TenChiPhi = "Chi phí lập HSMT và đánh giá HSDT thi công xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv6",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.21 - 2.23 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.7",
                    MaChiPhi = "TV_GS_XD",
                    TenChiPhi = "Chi phí giám sát thi công xây dựng",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gtv7",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.24 Thông tư 38/2026",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "5.8",
                    MaChiPhi = "TV_GS_TB",
                    TenChiPhi = "Chi phí giám sát lắp đặt thiết bị",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    KyHieu = "Gtv8",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Bảng 2.25 Thông tư 38/2026",
                    IsActive = false // Bật khi có chi phí thiết bị
                },

                // VI. Chi phí khác
                new ChiPhiKinhPhiItem
                {
                    STT = "6.1",
                    MaChiPhi = "K_TD_DA",
                    TenChiPhi = "Phí thẩm định dự án đầu tư xây dựng (hoặc Báo cáo KT-KT)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TongMucDauTu,
                    KyHieu = "Gk1",
                    ThueSuatGTGT = 0m, // Phí theo Luật Phí và Lệ phí
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định dự án",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.2",
                    MaChiPhi = "K_TD_TK",
                    TenChiPhi = "Phí thẩm định thiết kế xây dựng (sau TKCS)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gk2",
                    ThueSuatGTGT = 0m,
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định thiết kế",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.3",
                    MaChiPhi = "K_TD_DT",
                    TenChiPhi = "Phí thẩm định dự toán xây dựng",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    KyHieu = "Gk3",
                    ThueSuatGTGT = 0m,
                    MinValue = 500000m,
                    MaxValue = 150000000m,
                    GhiChuCachTinh = "Thông tư BTC về phí thẩm định dự toán",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.4",
                    MaChiPhi = "K_BH",
                    TenChiPhi = "Chi phí bảo hiểm công trình xây dựng",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    TyLePhanTram = 0.15m, // 0.15% cho dân dụng/giao thông
                    KyHieu = "Gk4",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Nghị định về bảo hiểm công trình",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.5",
                    MaChiPhi = "K_KT_DOCLAP",
                    TenChiPhi = "Chi phí kiểm toán độc lập",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gk5",
                    ThueSuatGTGT = 0.10m,
                    MinValue = 1000000m,
                    GhiChuCachTinh = "Quy định BTC về định mức chi phí kiểm toán độc lập",
                    IsActive = true
                },
                new ChiPhiKinhPhiItem
                {
                    STT = "6.6",
                    MaChiPhi = "K_TT_QUYETTOAN",
                    TenChiPhi = "Chi phí thẩm tra, phê duyệt quyết toán",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    KyHieu = "Gk6",
                    ThueSuatGTGT = 0m, // Phí ngân sách nhà nước
                    MinValue = 500000m,
                    GhiChuCachTinh = "Quy định BTC về định mức phí thẩm tra, phê duyệt quyết toán",
                    IsActive = true
                },

                // VII. Chi phí dự phòng
                new ChiPhiKinhPhiItem
                {
                    STT = "VII",
                    MaChiPhi = "G_DP",
                    TenChiPhi = "Chi phí dự phòng (Yếu tố khối lượng phát sinh)",
                    Nhom = NhomChiPhi.ChiPhiDuPhong,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.TongChiPhiTruocDuPhong,
                    TyLePhanTram = 5.0m, // 5% cho THDT (Bảng 2.1), 10% cho TMĐT (Bảng 1.2)
                    KyHieu = "Gdp",
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Khoản 1 Điều 3 TT 36/2026/TT-BXD",
                    IsActive = true,
                    IsReadOnly = true
                }
            };

            return list;
        }

        /// <summary>
        /// Cập nhật lại toàn bộ tỷ lệ % định mức, hệ số điều chỉnh và tính toán lại giá trị
        /// </summary>
        public static void CapNhatToanBoDinhMucVaTinhToan(BangTongHopKinhPhiModel model)
        {
            if (model == null) return;

            string loaiCT = model.LoaiCongTrinh;
            string capCT = model.CapCongTrinh;
            int soBuocTK = model.SoBuocThietKe;

            decimal gXD = model.ChiPhiXDTruocThue;
            decimal gTB = model.ChiPhiTBTruocThue;
            decimal gXDTB = gXD + gTB;

            // Chuyển sang đơn vị tỷ đồng để tra định mức
            decimal qmXDTy = gXD / 1_000_000_000m;
            decimal qmXDTBTy = gXDTB / 1_000_000_000m;
            if (qmXDTy <= 0) qmXDTy = 10m;     // Mặc định mốc tối thiểu 10 tỷ
            if (qmXDTBTy <= 0) qmXDTBTy = 10m;

            decimal heSoCap = DinhMucTT38Database.GetHeSoCapCongTrinh(capCT);

            // Hệ số điều chỉnh chung
            decimal kQLDA = 1.0m;
            if (model.CdtTuQuanLy) kQLDA *= 0.8m;
            if (model.VungKhoKhan) kQLDA *= 1.35m;
            if (model.TuyenQuaNhieuTinh) kQLDA *= 1.1m;

            decimal kTV = 1.0m;
            if (model.VungKhoKhan) kTV *= 1.35m;

            decimal kThietKe = kTV;
            if (model.CaiTaoSuaChua) kThietKe *= 1.15m;
            if (model.ThietKeLapLai) kThietKe *= 0.36m;

            // Hệ số điều chỉnh kiểm toán & quyết toán
            decimal kKiemToan = 1.0m;
            if (model.ThietBiTren50Pct) kKiemToan *= 0.7m;

            decimal kQuyetToan = 1.0m;
            if (model.ThietBiTren50Pct) kQuyetToan *= 0.7m;
            if (model.DaKiemToanDocLap) kQuyetToan *= 0.5m;

            // Hệ số thẩm định
            decimal kThamDinh = 1.0m;
            if (model.YeuCauThueThamTra) kThamDinh *= 0.5m;

            foreach (var item in model.Items)
            {
                if (item.CachTinh != CachTinhChiPhi.TheoTyLeDinhMuc)
                    continue;

                switch (item.MaChiPhi)
                {
                    case "G_QLDA":
                        if (DinhMucTT38Database.TiLeQLDA.TryGetValue(loaiCT, out var qldaArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoQLDA, qldaArr, qmXDTBTy);
                            item.HeSoDieuChinh = kQLDA;
                        }
                        break;

                    case "TV_FS_KTKT":
                        if (soBuocTK == 1 || qmXDTBTy < 15m)
                        {
                            // Báo cáo KT-KT
                            if (DinhMucTT38Database.TiLeBaoCaoKTKT.TryGetValue(loaiCT, out var ktktArr))
                            {
                                item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoBaoCaoKTKT, ktktArr, qmXDTBTy);
                                item.HeSoDieuChinh = kTV;
                                item.TenChiPhi = "Chi phí lập Báo cáo kinh tế - kỹ thuật";
                            }
                        }
                        else
                        {
                            // Lập FS
                            if (DinhMucTT38Database.TiLeLapFS.TryGetValue(loaiCT, out var fsArr))
                            {
                                item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoFS, fsArr, qmXDTBTy);
                                item.HeSoDieuChinh = kTV;
                                item.TenChiPhi = "Chi phí lập Báo cáo nghiên cứu khả thi (FS)";
                            }
                        }
                        break;

                    case "TV_TK":
                        if (DinhMucTT38Database.TiLeThietKeCap3.TryGetValue(loaiCT, out var tkArr))
                        {
                            decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThietKe, tkArr, qmXDTy);
                            item.TyLePhanTram = Math.Round(tlBase * heSoCap, 4);
                            item.HeSoDieuChinh = kThietKe;
                        }
                        break;

                    case "TV_TT_TK":
                        if (DinhMucTT38Database.TiLeThamTraTK.TryGetValue(loaiCT, out var tttkArr))
                        {
                            decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThamTraTK, tttkArr, qmXDTy);
                            item.TyLePhanTram = Math.Round(tlBase * heSoCap, 4);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_TT_DT":
                        if (DinhMucTT38Database.TiLeThamTraDuToan.TryGetValue(loaiCT, out var ttdtArr))
                        {
                            decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThamTraTK, ttdtArr, qmXDTy);
                            item.TyLePhanTram = Math.Round(tlBase * heSoCap, 4);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_HSMT":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoHSMT, DinhMucTT38Database.TiLeLapHSMT, qmXDTy);
                        item.HeSoDieuChinh = kTV;
                        break;

                    case "TV_GS_XD":
                        if (DinhMucTT38Database.TiLeGiamSatThiCong.TryGetValue(loaiCT, out var gsArr))
                        {
                            decimal tlBase = DinhMucChiPhiKhacBTC.NoiSuy(DinhMucTT38Database.MocQuyMoThietKe, gsArr, qmXDTy);
                            item.TyLePhanTram = Math.Round(tlBase * heSoCap, 4);
                            item.HeSoDieuChinh = kTV;
                        }
                        break;

                    case "TV_GS_TB":
                        item.TyLePhanTram = 0.65m;
                        item.HeSoDieuChinh = kTV;
                        if (gTB > 0) item.IsActive = true;
                        break;

                    case "K_TD_DA":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                            DinhMucChiPhiKhacBTC.MocQuyMoThamDinhDuAn,
                            DinhMucChiPhiKhacBTC.TiLeThamDinhDuAn,
                            qmXDTBTy);
                        item.HeSoDieuChinh = kThamDinh;
                        break;

                    case "K_TD_TK":
                        if (DinhMucChiPhiKhacBTC.TiLeThamDinhThietKe.TryGetValue(loaiCT, out var tdtkArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                                DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT,
                                tdtkArr,
                                qmXDTy);
                            item.HeSoDieuChinh = kThamDinh;
                        }
                        break;

                    case "K_TD_DT":
                        if (DinhMucChiPhiKhacBTC.TiLeThamDinhDuToan.TryGetValue(loaiCT, out var tddtArr))
                        {
                            item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                                DinhMucChiPhiKhacBTC.MocQuyMoThamDinhTKDT,
                                tddtArr,
                                qmXDTy);
                            item.HeSoDieuChinh = kThamDinh;
                        }
                        break;

                    case "K_BH":
                        item.TyLePhanTram = loaiCT.Contains("Giao thông") ? 0.15m : 0.10m;
                        break;

                    case "K_KT_DOCLAP":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                            DinhMucChiPhiKhacBTC.MocQuyMoKiemToan,
                            DinhMucChiPhiKhacBTC.TiLeKiemToanDocLap,
                            qmXDTBTy);
                        item.HeSoDieuChinh = kKiemToan;
                        break;

                    case "K_TT_QUYETTOAN":
                        item.TyLePhanTram = DinhMucChiPhiKhacBTC.NoiSuy(
                            DinhMucChiPhiKhacBTC.MocQuyMoQuyetToan,
                            DinhMucChiPhiKhacBTC.TiLeThamTraQuyetToan,
                            qmXDTBTy);
                        item.HeSoDieuChinh = kQuyetToan;
                        break;

                    case "G_DP":
                        // Giữ nguyên tỷ lệ người dùng cài (5% cho THDT, 10% cho TMĐT)
                        break;
                }
            }

            model.TinhToanLai();
        }

        /// <summary>
        /// Thư viện các khoản mục chi phí chuẩn để người dùng lựa chọn bổ sung vào dự toán
        /// </summary>
        public static List<ChiPhiKinhPhiItem> LayThuVienChiPhiChuan()
        {
            return new List<ChiPhiKinhPhiItem>
            {
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_KHAO_SAT_DH",
                    TenChiPhi = "Chi phí khảo sát địa hình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán khảo sát địa hình riêng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_KHAO_SAT_DC",
                    TenChiPhi = "Chi phí khảo sát địa chất công trình",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán khảo sát địa chất riêng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_DTM",
                    TenChiPhi = "Chi phí đánh giá tác động môi trường (ĐTM)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXD_Va_ThietBi,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Lập báo cáo ĐTM theo quy định Luật Môi trường"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_THAM_DINH_GIA_TB",
                    TenChiPhi = "Chi phí thẩm định giá thiết bị",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.TheoTyLeDinhMuc,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    TyLePhanTram = 0.25m,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Chứng thư thẩm định giá thiết bị"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_QUAN_TRAC",
                    TenChiPhi = "Chi phí quan trắc biến dạng công trình (lún, nghiêng)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Phương án quan trắc biến dạng công trình"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "TV_THI_NGHIEM_CHUYEN_NGANH",
                    TenChiPhi = "Chi phí thí nghiệm chuyên ngành xây dựng (nén cọc, siêu âm bê tông)",
                    Nhom = NhomChiPhi.TuVanDauTuXD,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Đề cương thí nghiệm kiểm định chất lượng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_RA_PHA_BOM_MIN",
                    TenChiPhi = "Chi phí rà phá bom mìn, vật nổ",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0m,
                    GhiChuCachTinh = "Dự toán rà phá bom mìn do Quân khu/Bộ CHQS phê duyệt"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_KHOI_CONG_KHANH_THANH",
                    TenChiPhi = "Chi phí tổ chức khởi công, khánh thành công trình",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Theo dự toán được cấp có thẩm quyền phê duyệt"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_DI_DOI_HA_TANG",
                    TenChiPhi = "Chi phí di dời các công trình hạ tầng kỹ thuật (điện, nước, viễn thông)",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiXayDung,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Dự toán di dời hoàn trả hạ tầng"
                },
                new ChiPhiKinhPhiItem
                {
                    MaChiPhi = "K_DANG_KIEM_AN_TOAN",
                    TenChiPhi = "Chi phí đăng kiểm an toàn thiết bị, PCCC",
                    Nhom = NhomChiPhi.ChiPhiKhac,
                    CachTinh = CachTinhChiPhi.NhapTruocThue,
                    CoSoTinh = CoSoTinhChiPhi.ChiPhiThietBi,
                    ThueSuatGTGT = 0.10m,
                    GhiChuCachTinh = "Kiểm định an toàn và nghiệm thu PCCC"
                }
            };
        }
    }
}
