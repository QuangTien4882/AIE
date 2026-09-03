using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models;

public enum PhamViBocXep
{
    CaHai,    // Cả bốc lên và bốc xuống
    BocLen,   // Chỉ bốc lên
    BocXuong  // Chỉ bốc xuống
}

public class DinhMucBocXepItem
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string LoaiVatLieu { get; set; } = string.Empty;
    public string DonViDinhMuc { get; set; } = string.Empty;
    public string PhuongPhap { get; set; } = "Thủ công"; // "Thủ công" hoặc "Cần cẩu 6t"
    
    // Hao phí nhân công nhóm I (công/ĐVT)
    public decimal DmNCLen { get; set; }
    public decimal DmNCXuong { get; set; }
    public decimal DmNCCaHai { get; set; }

    // Hao phí máy thi công (ca/ĐVT) nếu bốc xếp bằng cơ giới/cần cẩu
    public string? MaMay { get; set; }
    public string? TenMay { get; set; }
    public decimal DmMayLen { get; set; }
    public decimal DmMayXuong { get; set; }
    public decimal DmMayCaHai { get; set; }

    public bool CoMay => !string.IsNullOrEmpty(TenMay);

    public override string ToString() => $"{MaHieu} - {TenCongTac} ({DonViDinhMuc})";
}

public static class DinhMucBocXepDatabase
{
    public static readonly List<DinhMucBocXepItem> DanhSach = new()
    {
        // 1. Bốc xếp vật liệu rời lên phương tiện vận chuyển bằng thủ công (AM.11100) - ĐVT: m³
        new DinhMucBocXepItem
        {
            MaHieu = "AM.11101",
            TenCongTac = "Bốc xếp cát các loại lên phương tiện bằng thủ công",
            LoaiVatLieu = "Cát các loại",
            DonViDinhMuc = "m³",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.14m,
            DmNCXuong = 0m,
            DmNCCaHai = 0.14m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.11102",
            TenCongTac = "Bốc xếp đất các loại lên phương tiện bằng thủ công",
            LoaiVatLieu = "Đất các loại",
            DonViDinhMuc = "m³",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.17m,
            DmNCXuong = 0m,
            DmNCCaHai = 0.17m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.11103",
            TenCongTac = "Bốc xếp sỏi, đá dăm các loại lên phương tiện bằng thủ công",
            LoaiVatLieu = "Sỏi, đá dăm các loại",
            DonViDinhMuc = "m³",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.24m,
            DmNCXuong = 0m,
            DmNCCaHai = 0.24m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.11104",
            TenCongTac = "Bốc xếp đá hộc lên phương tiện bằng thủ công",
            LoaiVatLieu = "Đá hộc",
            DonViDinhMuc = "m³",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.33m,
            DmNCXuong = 0m,
            DmNCCaHai = 0.33m
        },

        // 2. Bốc lên, bốc xuống bằng thủ công (AM.11200)
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1121",
            TenCongTac = "Bốc lên, bốc xuống gạch xây các loại bằng thủ công",
            LoaiVatLieu = "Gạch xây các loại",
            DonViDinhMuc = "1000 viên",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.34m,
            DmNCXuong = 0.31m,
            DmNCCaHai = 0.65m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1122",
            TenCongTac = "Bốc lên, bốc xuống gạch ốp, lát các loại bằng thủ công",
            LoaiVatLieu = "Gạch ốp, lát các loại",
            DonViDinhMuc = "1000 viên",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.20m,
            DmNCXuong = 0.19m,
            DmNCCaHai = 0.39m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1123",
            TenCongTac = "Bốc lên, bốc xuống ngói các loại bằng thủ công",
            LoaiVatLieu = "Ngói các loại",
            DonViDinhMuc = "1000 viên",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.38m,
            DmNCXuong = 0.38m,
            DmNCCaHai = 0.76m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1124",
            TenCongTac = "Bốc lên, bốc xuống xi măng bao bằng thủ công",
            LoaiVatLieu = "Xi măng bao",
            DonViDinhMuc = "tấn",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.15m,
            DmNCXuong = 0.11m,
            DmNCCaHai = 0.26m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1125",
            TenCongTac = "Bốc lên, bốc xuống gỗ các loại bằng thủ công",
            LoaiVatLieu = "Gỗ các loại",
            DonViDinhMuc = "m³",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.14m,
            DmNCXuong = 0.09m,
            DmNCCaHai = 0.23m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1126",
            TenCongTac = "Bốc lên, bốc xuống cọc gỗ, cừ tràm bằng thủ công",
            LoaiVatLieu = "Cọc gỗ, cừ tràm",
            DonViDinhMuc = "100 cây",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.87m,
            DmNCXuong = 0.56m,
            DmNCCaHai = 1.43m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1127",
            TenCongTac = "Bốc lên, bốc xuống tre, cây chống bằng thủ công",
            LoaiVatLieu = "Tre, cây chống",
            DonViDinhMuc = "100 cây",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.75m,
            DmNCXuong = 0.47m,
            DmNCCaHai = 1.22m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.1128",
            TenCongTac = "Bốc lên, bốc xuống thép các loại bằng thủ công",
            LoaiVatLieu = "Thép các loại",
            DonViDinhMuc = "tấn",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.34m,
            DmNCXuong = 0.21m,
            DmNCCaHai = 0.55m
        },

        // 3. Bốc xếp cấu kiện bê tông đúc sẵn bằng thủ công (AM.11600)
        new DinhMucBocXepItem
        {
            MaHieu = "AM.116",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng P ≤ 200kg bằng thủ công",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 200kg (thủ công)",
            DonViDinhMuc = "tấn",
            PhuongPhap = "Thủ công",
            DmNCLen = 0.20m,
            DmNCXuong = 0.13m,
            DmNCCaHai = 0.33m
        },

        // 4. Bốc xếp cấu kiện bê tông đúc sẵn bằng cần cẩu (AM.12000)
        new DinhMucBocXepItem
        {
            MaHieu = "AM.121",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng P ≤ 200kg bằng cần cẩu",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 200kg (cần cẩu)",
            DonViDinhMuc = "cấu kiện",
            PhuongPhap = "Cần cẩu 6t",
            DmNCLen = 0.030m,
            DmNCXuong = 0.022m,
            DmNCCaHai = 0.052m,
            MaMay = "M102.0201",
            TenMay = "Cần cẩu bánh hơi - sức nâng: 6 T",
            DmMayLen = 0.014m,
            DmMayXuong = 0.011m,
            DmMayCaHai = 0.025m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.122",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng ≤ 500kg bằng cần cẩu",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 500kg (cần cẩu)",
            DonViDinhMuc = "cấu kiện",
            PhuongPhap = "Cần cẩu 6t",
            DmNCLen = 0.06m,
            DmNCXuong = 0.05m,
            DmNCCaHai = 0.11m,
            MaMay = "M102.0201",
            TenMay = "Cần cẩu bánh hơi - sức nâng: 6 T",
            DmMayLen = 0.020m,
            DmMayXuong = 0.016m,
            DmMayCaHai = 0.036m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.123",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng ≤ 1T bằng cần cẩu",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 1T (cần cẩu)",
            DonViDinhMuc = "cấu kiện",
            PhuongPhap = "Cần cẩu 6t",
            DmNCLen = 0.08m,
            DmNCXuong = 0.06m,
            DmNCCaHai = 0.14m,
            MaMay = "M102.0201",
            TenMay = "Cần cẩu bánh hơi - sức nâng: 6 T",
            DmMayLen = 0.026m,
            DmMayXuong = 0.020m,
            DmMayCaHai = 0.046m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.124",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng ≤ 2T bằng cần cẩu",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 2T (cần cẩu)",
            DonViDinhMuc = "cấu kiện",
            PhuongPhap = "Cần cẩu 6t",
            DmNCLen = 0.09m,
            DmNCXuong = 0.08m,
            DmNCCaHai = 0.17m,
            MaMay = "M102.0201",
            TenMay = "Cần cẩu bánh hơi - sức nâng: 6 T",
            DmMayLen = 0.030m,
            DmMayXuong = 0.024m,
            DmMayCaHai = 0.054m
        },
        new DinhMucBocXepItem
        {
            MaHieu = "AM.125",
            TenCongTac = "Bốc xếp cấu kiện bê tông đúc sẵn trọng lượng ≤ 5T bằng cần cẩu",
            LoaiVatLieu = "Cấu kiện BT đúc sẵn P ≤ 5T (cần cẩu)",
            DonViDinhMuc = "cấu kiện",
            PhuongPhap = "Cần cẩu 6t",
            DmNCLen = 0.13m,
            DmNCXuong = 0.11m,
            DmNCCaHai = 0.24m,
            MaMay = "M102.0201",
            TenMay = "Cần cẩu bánh hơi - sức nâng: 6 T",
            DmMayLen = 0.043m,
            DmMayXuong = 0.034m,
            DmMayCaHai = 0.077m
        }
    };

    /// <summary>
    /// Nhận diện định mức bốc xếp phù hợp dựa trên tên và đơn vị vật liệu.
    /// Trả về null nếu vật liệu không có định mức bốc xếp theo TT12.
    /// </summary>
    public static DinhMucBocXepItem? NhanDienDinhMuc(string tenVatLieu, string donVi = "")
    {
        if (string.IsNullOrWhiteSpace(tenVatLieu)) return null;

        string t = tenVatLieu.Trim().ToLower();
        string dv = (donVi ?? "").Trim().ToLower();

        // Cát
        if (t.Contains("cát") && !t.Contains("cát xét") && !t.Contains("que"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.11101");
        }

        // Đất
        if (t.Contains("đất") && !t.Contains("đất nung") && !t.Contains("đất sét"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.11102");
        }

        // Đá hộc
        if (t.Contains("đá hộc"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.11104");
        }

        // Sỏi, đá dăm các loại (đá 1x2, đá 2x4, đá 4x6, đá mi...)
        if (t.Contains("sỏi") || t.Contains("đá dăm") || t.Contains("đá 1x2") || t.Contains("đá 2x4") ||
            t.Contains("đá 4x6") || t.Contains("đá 0.5x1") || t.Contains("đá mi") || t.Contains("đá mạt") ||
            (t.StartsWith("đá ") && (dv == "m3" || dv == "m³")))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.11103");
        }

        // Gạch ốp, lát
        if (t.Contains("ốp") || t.Contains("lát") || t.Contains("ceramic") || t.Contains("granite") || t.Contains("gạch men") || t.Contains("len chân tường"))
        {
            if (t.Contains("gạch") || dv.Contains("viên") || dv.Contains("v") || dv.Contains("m2") || dv.Contains("m²"))
                return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1122");
        }

        // Gạch xây các loại
        if (t.Contains("gạch") && (t.Contains("xây") || t.Contains("ống") || t.Contains("thẻ") || t.Contains("đặc") || t.Contains("lỗ") || t.Contains("block") || t.Contains("tuynel") || t.Contains("bê tông")))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1121");
        }
        if (t.StartsWith("gạch ") && !t.Contains("gạch vỡ"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1121");
        }

        // Ngói các loại
        if (t.Contains("ngói"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1123");
        }

        // Xi măng bao
        if (t.Contains("xi măng") || t.Contains("pcb") || t.Contains("pc30") || t.Contains("pc40"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1124");
        }

        // Gỗ các loại
        if (t.Contains("gỗ") && !t.Contains("cọc gỗ") && (dv == "m3" || dv == "m³" || t.Contains("ván") || t.Contains("xà gồ") || t.Contains("khuôn")))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1125");
        }

        // Cọc gỗ, cừ tràm
        if (t.Contains("cừ tràm") || t.Contains("cọc gỗ") || t.Contains("cừ bạch đàn"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1126");
        }

        // Tre, cây chống
        if (t.Contains("tre") || t.Contains("cây chống") || t.Contains("luồng") || t.Contains("nứa"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1127");
        }

        // Thép các loại
        if (t.Contains("thép") || t.Contains("sắt") || t.Contains("cốt thép") || t.Contains("tôn") || t.Contains("lưới b40"))
        {
            // Tránh que hàn, đinh thép
            if (!t.Contains("que hàn") && !t.Contains("đinh") && !t.Contains("bulông") && !t.Contains("bu lông") && !t.Contains("vít"))
                return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.1128");
        }

        // Cấu kiện bê tông đúc sẵn
        if (t.Contains("cấu kiện bê tông") || t.Contains("bê tông đúc sẵn") || t.Contains("ống cống") || t.Contains("tấm đan") || t.Contains("cọc bê tông"))
        {
            return DanhSach.FirstOrDefault(x => x.MaHieu == "AM.116");
        }

        return null;
    }

    /// <summary>
    /// Hệ số quy đổi từ 1 ĐVT của định mức sang 1 ĐVT của vật tư.
    /// Ví dụ: Gạch định mức tính cho "1000 viên", vật tư tính cho "viên" => hệ số quy đổi = 1/1000 = 0.001
    /// Thép định mức tính cho "tấn", vật tư tính cho "kg" => hệ số quy đổi = 1/1000 = 0.001
    /// Cừ tràm định mức tính cho "100 cây", vật tư tính cho "cây" => hệ số quy đổi = 1/100 = 0.01
    /// </summary>
    public static decimal TinhHeSoQuyDoi(string donViVatTu, string donViDinhMuc)
    {
        string dvVT = (donViVatTu ?? "").Trim().ToLower();
        string dvDM = (donViDinhMuc ?? "").Trim().ToLower();

        if (dvVT == dvDM) return 1.0m;

        // 1000 viên -> 1 viên
        if ((dvDM.Contains("1000") || dvDM.Contains("1.000") || dvDM.Contains("nghìn")) && 
            (dvVT == "viên" || dvVT == "v" || dvVT == "vien"))
        {
            return 0.001m;
        }

        // tấn -> kg
        if ((dvDM == "tấn" || dvDM == "tan" || dvDM == "t") && 
            (dvVT == "kg" || dvVT == "kí" || dvVT == "kilogam"))
        {
            return 0.001m;
        }

        // tấn -> tạ
        if ((dvDM == "tấn" || dvDM == "tan" || dvDM == "t") && dvVT == "tạ")
        {
            return 0.1m;
        }

        // 100 cây -> 1 cây
        if ((dvDM.Contains("100") || dvDM.Contains("trăm")) && 
            (dvVT == "cây" || dvVT == "cay" || dvVT == "cọc" || dvVT == "coc"))
        {
            return 0.01m;
        }

        // tấn -> bao (xi măng 50kg)
        if ((dvDM == "tấn" || dvDM == "tan") && dvVT == "bao")
        {
            return 0.05m; // 1 tấn = 20 bao 50kg -> 1 bao = 0.05 tấn
        }

        return 1.0m;
    }
}
