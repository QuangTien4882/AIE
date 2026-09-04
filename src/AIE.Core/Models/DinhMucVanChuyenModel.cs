using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Core.Models;

/// <summary>
/// Định mức vận chuyển bằng ô tô theo Chương XII - Thông tư số 12/2021/TT-BXD & TT 38/2026/TT-BXD
/// </summary>
public class DinhMucVCOToItem
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string LoaiVatLieu { get; set; } = string.Empty;
    public string LoaiXe { get; set; } = string.Empty; // "Ô tô tự đổ" hoặc "Ô tô vận tải thùng"
    public decimal TaiTrong { get; set; } // 7, 10, 12, 20, 22
    public string MaMay { get; set; } = string.Empty; // M108.0102, etc.
    public string TenMay { get; set; } = string.Empty;
    public string DonViDinhMuc { get; set; } = string.Empty; // 10m3 hoặc 10tấn
    
    public decimal Dm1 { get; set; } // Cự ly <= 1km
    public decimal Dm2 { get; set; } // 1km tiếp theo trong phạm vi <= 10km
    public decimal Dm3 { get; set; } // 1km tiếp theo trong phạm vi <= 60km

    public override string ToString() => $"{MaHieu} - {TenCongTac} ({TenMay})";
}

/// <summary>
/// Định mức vận chuyển bằng thủ công (bộ) theo AM.21000
/// </summary>
public class DinhMucVCBoItem
{
    public string MaHieu { get; set; } = string.Empty;
    public string TenCongTac { get; set; } = string.Empty;
    public string LoaiVatLieu { get; set; } = string.Empty;
    public string DonViDinhMuc { get; set; } = string.Empty;
    
    public decimal Dm10m { get; set; } // 10m khởi điểm
    public decimal DmTiepTheo { get; set; } // Mỗi 10m tiếp theo trong phạm vi <= 300m

    public override string ToString() => $"{MaHieu} - {TenCongTac} ({DonViDinhMuc})";
}

/// <summary>
/// Cung đoạn vận chuyển ô tô
/// </summary>
public class CungDuongVanChuyen
{
    public string NacCuLy { get; set; } = string.Empty;
    public string DiemDau { get; set; } = string.Empty;
    public string DiemCuoi { get; set; } = string.Empty;
    public string TenDoanDuong { get; set; } = string.Empty;
    public decimal CuLyKm { get; set; }
    public int LoaiDuong { get; set; } = 3; // Mặc định loại 3

    public decimal HeSoK => LoaiDuong switch
    {
        1 => 0.57m,
        2 => 0.68m,
        3 => 1.00m,
        4 => 1.35m,
        5 => 1.50m,
        6 => 1.80m,
        _ => 1.00m
    };
}

/// <summary>
/// Đoạn vận chuyển bộ thủ công
/// </summary>
public class DoanVanChuyenBo
{
    public string NacCuLy { get; set; } = string.Empty;
    public string DiemDau { get; set; } = string.Empty;
    public string DiemCuoi { get; set; } = string.Empty;
    public decimal CuLyMet { get; set; }
    public decimal HeSoDiaHinh { get; set; } = 1.0m;
    public int SoTang { get; set; } = 1;
}

/// <summary>
/// Tuyến đường theo Quyết định số 2035/QĐ-UBND TP Đà Nẵng
/// </summary>
public class PhanLoaiDuongItem
{
    public string TuyenDuong { get; set; } = string.Empty;
    public string DiaPhan { get; set; } = string.Empty;
    public string TuKmDenKm { get; set; } = string.Empty;
    public decimal ChieuDaiKm { get; set; }
    public int LoaiDuong { get; set; } = 3;
    public string GhiChu { get; set; } = string.Empty;

    public override string ToString() => $"{TuyenDuong} ({TuKmDenKm}) - {ChieuDaiKm:0.###}km [Loại {LoaiDuong}]";
}

public static class DinhMucVanChuyenDatabase
{
    // ==========================================
    // 1. DANH SÁCH ĐỊNH MỨC VẬN CHUYỂN Ô TÔ
    // ==========================================
    public static readonly List<DinhMucVCOToItem> DanhSachOTo = new()
    {
        // AM.23100: Vận chuyển cát bằng ô tô tự đổ (10m3)
        new() { MaHieu = "AM.2311", TenCongTac = "Vận chuyển cát bằng ôtô tự đổ 7 T", LoaiVatLieu = "Cát các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 7, MaMay = "M106.0203", TenMay = "Ô tô tự đổ - trọng tải: 7 T", DonViDinhMuc = "10m³", Dm1 = 0.027m, Dm2 = 0.019m, Dm3 = 0.014m },
        new() { MaHieu = "AM.2312", TenCongTac = "Vận chuyển cát bằng ôtô tự đổ 10 T", LoaiVatLieu = "Cát các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 10, MaMay = "M106.0204", TenMay = "Ô tô tự đổ - trọng tải: 10 T", DonViDinhMuc = "10m³", Dm1 = 0.020m, Dm2 = 0.015m, Dm3 = 0.010m },
        new() { MaHieu = "AM.2313", TenCongTac = "Vận chuyển cát bằng ôtô tự đổ 12 T", LoaiVatLieu = "Cát các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 12, MaMay = "M106.0205", TenMay = "Ô tô tự đổ - trọng tải: 12 T", DonViDinhMuc = "10m³", Dm1 = 0.016m, Dm2 = 0.012m, Dm3 = 0.008m },
        new() { MaHieu = "AM.2314", TenCongTac = "Vận chuyển cát bằng ôtô tự đổ 22 T", LoaiVatLieu = "Cát các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 22, MaMay = "M106.0208", TenMay = "Ô tô tự đổ - trọng tải: 22 T", DonViDinhMuc = "10m³", Dm1 = 0.011m, Dm2 = 0.008m, Dm3 = 0.004m },

        // AM.23200: Vận chuyển đất bằng ô tô tự đổ (10m3)
        new() { MaHieu = "AM.2321", TenCongTac = "Vận chuyển đất bằng ôtô tự đổ 7 T", LoaiVatLieu = "Đất các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 7, MaMay = "M106.0203", TenMay = "Ô tô tự đổ - trọng tải: 7 T", DonViDinhMuc = "10m³", Dm1 = 0.030m, Dm2 = 0.021m, Dm3 = 0.015m },
        new() { MaHieu = "AM.2322", TenCongTac = "Vận chuyển đất bằng ôtô tự đổ 10 T", LoaiVatLieu = "Đất các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 10, MaMay = "M106.0204", TenMay = "Ô tô tự đổ - trọng tải: 10 T", DonViDinhMuc = "10m³", Dm1 = 0.022m, Dm2 = 0.016m, Dm3 = 0.011m },
        new() { MaHieu = "AM.2323", TenCongTac = "Vận chuyển đất bằng ôtô tự đổ 12 T", LoaiVatLieu = "Đất các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 12, MaMay = "M106.0205", TenMay = "Ô tô tự đổ - trọng tải: 12 T", DonViDinhMuc = "10m³", Dm1 = 0.018m, Dm2 = 0.013m, Dm3 = 0.009m },
        new() { MaHieu = "AM.2324", TenCongTac = "Vận chuyển đất bằng ôtô tự đổ 22 T", LoaiVatLieu = "Đất các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 22, MaMay = "M106.0208", TenMay = "Ô tô tự đổ - trọng tải: 22 T", DonViDinhMuc = "10m³", Dm1 = 0.012m, Dm2 = 0.008m, Dm3 = 0.005m },

        // AM.23400: Vận chuyển đá dăm bằng ô tô tự đổ (10m3)
        new() { MaHieu = "AM.2341", TenCongTac = "Vận chuyển đá dăm các loại bằng ôtô tự đổ 7 T", LoaiVatLieu = "Đá dăm các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 7, MaMay = "M106.0203", TenMay = "Ô tô tự đổ - trọng tải: 7 T", DonViDinhMuc = "10m³", Dm1 = 0.034m, Dm2 = 0.025m, Dm3 = 0.018m },
        new() { MaHieu = "AM.2342", TenCongTac = "Vận chuyển đá dăm các loại bằng ôtô tự đổ 10 T", LoaiVatLieu = "Đá dăm các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 10, MaMay = "M106.0204", TenMay = "Ô tô tự đổ - trọng tải: 10 T", DonViDinhMuc = "10m³", Dm1 = 0.026m, Dm2 = 0.019m, Dm3 = 0.013m },
        new() { MaHieu = "AM.2343", TenCongTac = "Vận chuyển đá dăm các loại bằng ôtô tự đổ 12 T", LoaiVatLieu = "Đá dăm các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 12, MaMay = "M106.0205", TenMay = "Ô tô tự đổ - trọng tải: 12 T", DonViDinhMuc = "10m³", Dm1 = 0.021m, Dm2 = 0.016m, Dm3 = 0.010m },
        new() { MaHieu = "AM.2344", TenCongTac = "Vận chuyển đá dăm các loại bằng ôtô tự đổ 22 T", LoaiVatLieu = "Đá dăm các loại", LoaiXe = "Ô tô tự đổ", TaiTrong = 22, MaMay = "M106.0208", TenMay = "Ô tô tự đổ - trọng tải: 22 T", DonViDinhMuc = "10m³", Dm1 = 0.014m, Dm2 = 0.009m, Dm3 = 0.007m },

        // AM.23500: Vận chuyển đá hộc bằng ô tô tự đổ (10m3)
        new() { MaHieu = "AM.2351", TenCongTac = "Vận chuyển đá hộc bằng ôtô tự đổ 7 T", LoaiVatLieu = "Đá hộc", LoaiXe = "Ô tô tự đổ", TaiTrong = 7, MaMay = "M106.0203", TenMay = "Ô tô tự đổ - trọng tải: 7 T", DonViDinhMuc = "10m³", Dm1 = 0.034m, Dm2 = 0.025m, Dm3 = 0.016m },
        new() { MaHieu = "AM.2352", TenCongTac = "Vận chuyển đá hộc bằng ôtô tự đổ 10 T", LoaiVatLieu = "Đá hộc", LoaiXe = "Ô tô tự đổ", TaiTrong = 10, MaMay = "M106.0204", TenMay = "Ô tô tự đổ - trọng tải: 10 T", DonViDinhMuc = "10m³", Dm1 = 0.025m, Dm2 = 0.018m, Dm3 = 0.012m },
        new() { MaHieu = "AM.2353", TenCongTac = "Vận chuyển đá hộc bằng ôtô tự đổ 12 T", LoaiVatLieu = "Đá hộc", LoaiXe = "Ô tô tự đổ", TaiTrong = 12, MaMay = "M106.0205", TenMay = "Ô tô tự đổ - trọng tải: 12 T", DonViDinhMuc = "10m³", Dm1 = 0.020m, Dm2 = 0.015m, Dm3 = 0.009m },
        new() { MaHieu = "AM.2354", TenCongTac = "Vận chuyển đá hộc bằng ôtô tự đổ 22 T", LoaiVatLieu = "Đá hộc", LoaiXe = "Ô tô tự đổ", TaiTrong = 22, MaMay = "M106.0208", TenMay = "Ô tô tự đổ - trọng tải: 22 T", DonViDinhMuc = "10m³", Dm1 = 0.013m, Dm2 = 0.009m, Dm3 = 0.006m },

        // AM.24100: Vận chuyển gạch xây bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2411", TenCongTac = "Vận chuyển gạch xây các loại bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Gạch xây các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.076m, Dm2 = 0.055m, Dm3 = 0.037m },
        new() { MaHieu = "AM.2412", TenCongTac = "Vận chuyển gạch xây các loại bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Gạch xây các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.049m, Dm2 = 0.036m, Dm3 = 0.023m },
        new() { MaHieu = "AM.2413", TenCongTac = "Vận chuyển gạch xây các loại bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Gạch xây các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.028m, Dm2 = 0.020m, Dm3 = 0.014m },

        // AM.24200: Vận chuyển gạch ốp lát bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2421", TenCongTac = "Vận chuyển gạch ốp lát các loại bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Gạch ốp lát các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.108m, Dm2 = 0.078m, Dm3 = 0.053m },
        new() { MaHieu = "AM.2422", TenCongTac = "Vận chuyển gạch ốp lát các loại bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Gạch ốp lát các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.072m, Dm2 = 0.051m, Dm3 = 0.035m },
        new() { MaHieu = "AM.2423", TenCongTac = "Vận chuyển gạch ốp lát các loại bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Gạch ốp lát các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.043m, Dm2 = 0.029m, Dm3 = 0.020m },

        // AM.24300: Vận chuyển ngói bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2431", TenCongTac = "Vận chuyển ngói các loại bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Ngói các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.090m, Dm2 = 0.066m, Dm3 = 0.045m },
        new() { MaHieu = "AM.2432", TenCongTac = "Vận chuyển ngói các loại bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Ngói các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.059m, Dm2 = 0.043m, Dm3 = 0.031m },
        new() { MaHieu = "AM.2433", TenCongTac = "Vận chuyển ngói các loại bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Ngói các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.033m, Dm2 = 0.024m, Dm3 = 0.017m },

        // AM.24400: Vận chuyển xi măng bao bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2441", TenCongTac = "Vận chuyển xi măng bao bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Xi măng bao", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.043m, Dm2 = 0.031m, Dm3 = 0.021m },
        new() { MaHieu = "AM.2442", TenCongTac = "Vận chuyển xi măng bao bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Xi măng bao", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.027m, Dm2 = 0.019m, Dm3 = 0.013m },
        new() { MaHieu = "AM.2443", TenCongTac = "Vận chuyển xi măng bao bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Xi măng bao", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.016m, Dm2 = 0.011m, Dm3 = 0.008m },

        // AM.24500: Vận chuyển thép các loại bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2451", TenCongTac = "Vận chuyển thép các loại bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Thép các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.022m, Dm2 = 0.016m, Dm3 = 0.011m },
        new() { MaHieu = "AM.2452", TenCongTac = "Vận chuyển thép các loại bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Thép các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.013m, Dm2 = 0.010m, Dm3 = 0.006m },
        new() { MaHieu = "AM.2453", TenCongTac = "Vận chuyển thép các loại bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Thép các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.007m, Dm2 = 0.006m, Dm3 = 0.003m },

        // AM.24600: Vận chuyển nhựa đường bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2461", TenCongTac = "Vận chuyển nhựa đường bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Nhựa đường", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.031m, Dm2 = 0.023m, Dm3 = 0.015m },
        new() { MaHieu = "AM.2462", TenCongTac = "Vận chuyển nhựa đường bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Nhựa đường", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.019m, Dm2 = 0.014m, Dm3 = 0.012m },
        new() { MaHieu = "AM.2463", TenCongTac = "Vận chuyển nhựa đường bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Nhựa đường", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.011m, Dm2 = 0.009m, Dm3 = 0.005m },

        // AM.24700: Vận chuyển gỗ các loại bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2471", TenCongTac = "Vận chuyển gỗ các loại bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Gỗ các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.024m, Dm2 = 0.018m, Dm3 = 0.011m },
        new() { MaHieu = "AM.2472", TenCongTac = "Vận chuyển gỗ các loại bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Gỗ các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.015m, Dm2 = 0.011m, Dm3 = 0.006m },
        new() { MaHieu = "AM.2473", TenCongTac = "Vận chuyển gỗ các loại bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Gỗ các loại", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.009m, Dm2 = 0.006m, Dm3 = 0.003m },

        // AM.25000: Vận chuyển cấu kiện bê tông đúc sẵn bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2511", TenCongTac = "Vận chuyển cấu kiện bê tông bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Cấu kiện bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.024m, Dm2 = 0.019m, Dm3 = 0.015m },
        new() { MaHieu = "AM.2512", TenCongTac = "Vận chuyển cấu kiện bê tông bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Cấu kiện bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.016m, Dm2 = 0.013m, Dm3 = 0.010m },
        new() { MaHieu = "AM.2513", TenCongTac = "Vận chuyển cấu kiện bê tông bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Cấu kiện bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.011m, Dm2 = 0.009m, Dm3 = 0.006m },

        // AM.26000: Vận chuyển ống cống bê tông bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2611", TenCongTac = "Vận chuyển ống cống bê tông bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Ống cống bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.026m, Dm2 = 0.021m, Dm3 = 0.017m },
        new() { MaHieu = "AM.2612", TenCongTac = "Vận chuyển ống cống bê tông bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Ống cống bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.018m, Dm2 = 0.015m, Dm3 = 0.012m },
        new() { MaHieu = "AM.2613", TenCongTac = "Vận chuyển ống cống bê tông bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Ống cống bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.013m, Dm2 = 0.011m, Dm3 = 0.009m },

        // AM.27000: Vận chuyển cọc, cột bê tông bằng ô tô thùng (10 tấn)
        new() { MaHieu = "AM.2711", TenCongTac = "Vận chuyển cọc, cột bê tông bằng ôtô vận tải thùng 7 T", LoaiVatLieu = "Cọc, cột bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 7, MaMay = "M106.0106", TenMay = "Ô tô vận tải thùng - trọng tải: 7 T", DonViDinhMuc = "10tấn", Dm1 = 0.025m, Dm2 = 0.020m, Dm3 = 0.016m },
        new() { MaHieu = "AM.2712", TenCongTac = "Vận chuyển cọc, cột bê tông bằng ôtô vận tải thùng 12 T", LoaiVatLieu = "Cọc, cột bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 12, MaMay = "M106.0108", TenMay = "Ô tô vận tải thùng - trọng tải: 12 T", DonViDinhMuc = "10tấn", Dm1 = 0.017m, Dm2 = 0.014m, Dm3 = 0.011m },
        new() { MaHieu = "AM.2713", TenCongTac = "Vận chuyển cọc, cột bê tông bằng ôtô vận tải thùng 20 T", LoaiVatLieu = "Cọc, cột bê tông", LoaiXe = "Ô tô vận tải thùng", TaiTrong = 20, MaMay = "M106.0110", TenMay = "Ô tô vận tải thùng - trọng tải: 20 T", DonViDinhMuc = "10tấn", Dm1 = 0.012m, Dm2 = 0.010m, Dm3 = 0.008m }
    };

    // ==========================================
    // 2. DANH SÁCH ĐỊNH MỨC VẬN CHUYỂN BỘ (THỦ CÔNG)
    // ==========================================
    public static readonly List<DinhMucVCBoItem> DanhSachBo = new()
    {
        new() { MaHieu = "AM.2101", TenCongTac = "Vận chuyển cát các loại bằng thủ công", LoaiVatLieu = "Cát các loại", DonViDinhMuc = "m³", Dm10m = 0.075m, DmTiepTheo = 0.008m },
        new() { MaHieu = "AM.2102", TenCongTac = "Vận chuyển đất các loại bằng thủ công", LoaiVatLieu = "Đất các loại", DonViDinhMuc = "m³", Dm10m = 0.088m, DmTiepTheo = 0.010m },
        new() { MaHieu = "AM.2103", TenCongTac = "Vận chuyển sỏi, đá dăm các loại bằng thủ công", LoaiVatLieu = "Sỏi, đá dăm các loại", DonViDinhMuc = "m³", Dm10m = 0.075m, DmTiepTheo = 0.009m },
        new() { MaHieu = "AM.2104", TenCongTac = "Vận chuyển đá hộc bằng thủ công", LoaiVatLieu = "Đá hộc", DonViDinhMuc = "m³", Dm10m = 0.088m, DmTiepTheo = 0.010m },
        new() { MaHieu = "AM.2105", TenCongTac = "Vận chuyển gạch xây các loại bằng thủ công", LoaiVatLieu = "Gạch xây các loại", DonViDinhMuc = "1000v", Dm10m = 0.075m, DmTiepTheo = 0.008m },
        new() { MaHieu = "AM.2106", TenCongTac = "Vận chuyển gạch ốp, lát các loại bằng thủ công", LoaiVatLieu = "Gạch ốp, lát các loại", DonViDinhMuc = "1000v", Dm10m = 0.038m, DmTiepTheo = 0.004m },
        new() { MaHieu = "AM.2107", TenCongTac = "Vận chuyển ngói các loại bằng thủ công", LoaiVatLieu = "Ngói các loại", DonViDinhMuc = "1000v", Dm10m = 0.090m, DmTiepTheo = 0.010m },
        new() { MaHieu = "AM.2108", TenCongTac = "Vận chuyển xi măng bao bằng thủ công", LoaiVatLieu = "Xi măng bao", DonViDinhMuc = "tấn", Dm10m = 0.075m, DmTiepTheo = 0.008m },
        new() { MaHieu = "AM.2109", TenCongTac = "Vận chuyển gỗ các loại bằng thủ công", LoaiVatLieu = "Gỗ các loại", DonViDinhMuc = "m³", Dm10m = 0.050m, DmTiepTheo = 0.006m },
        new() { MaHieu = "AM.2110", TenCongTac = "Vận chuyển cọc gỗ, cừ tràm bằng thủ công", LoaiVatLieu = "Cọc gỗ, cừ tràm", DonViDinhMuc = "100cây", Dm10m = 0.054m, DmTiepTheo = 0.006m },
        new() { MaHieu = "AM.2111", TenCongTac = "Vận chuyển tre, cây chống bằng thủ công", LoaiVatLieu = "Tre, cây chống", DonViDinhMuc = "100cây", Dm10m = 0.063m, DmTiepTheo = 0.007m },
        new() { MaHieu = "AM.2112", TenCongTac = "Vận chuyển sắt thép các loại bằng thủ công", LoaiVatLieu = "Sắt thép các loại", DonViDinhMuc = "tấn", Dm10m = 0.081m, DmTiepTheo = 0.009m }
    };

    // ==========================================
    // 3. DANH MỤC PHÂN LOẠI ĐƯỜNG ĐÀ NẴNG (QĐ 2035/QĐ-UBND)
    // ==========================================
    public static readonly List<PhanLoaiDuongItem> DanhSachDuongDaNang = new()
    {
        new() { TuyenDuong = "QL.14B", DiaPhan = "Sơn Trà - Hòa Vang", TuKmDenKm = "Km0+00 - Km24+100", ChieuDaiKm = 24.10m, LoaiDuong = 2, GhiChu = "Cảng Tiên Sa" },
        new() { TuyenDuong = "QL.14B", DiaPhan = "Hòa Vang", TuKmDenKm = "Km24+100 - Km32+126", ChieuDaiKm = 8.026m, LoaiDuong = 4, GhiChu = "Đoạn đang cải tạo" },
        new() { TuyenDuong = "QL.14B", DiaPhan = "Hòa Vang", TuKmDenKm = "Km32+126 - Km50+00", ChieuDaiKm = 17.90m, LoaiDuong = 2, GhiChu = "Giao đường HCM" },
        new() { TuyenDuong = "QL.14B", DiaPhan = "Đại Lộc", TuKmDenKm = "Km50+00 - Km73+971", ChieuDaiKm = 24.00m, LoaiDuong = 3, GhiChu = "Thạnh Mỹ" },
        new() { TuyenDuong = "QL.14D", DiaPhan = "Nam Giang", TuKmDenKm = "Km0+00 - Km74+387", ChieuDaiKm = 74.40m, LoaiDuong = 5, GhiChu = "Bến Giằng - CK Nam Giang" },
        new() { TuyenDuong = "QL.14E", DiaPhan = "Thăng Bình", TuKmDenKm = "Km0+00 - Km4+500", ChieuDaiKm = 4.50m, LoaiDuong = 4, GhiChu = "Giao ĐT.613B" },
        new() { TuyenDuong = "QL.14E", DiaPhan = "Thăng Bình", TuKmDenKm = "Km4+500 - Km9+060", ChieuDaiKm = 4.60m, LoaiDuong = 3, GhiChu = "Thăng Bình" },
        new() { TuyenDuong = "QL.14E", DiaPhan = "Thăng Bình", TuKmDenKm = "Km9+060 - Km11+00", ChieuDaiKm = 2.30m, LoaiDuong = 2, GhiChu = "Trùng QL.1" },
        new() { TuyenDuong = "QL.14E", DiaPhan = "Thăng Bình", TuKmDenKm = "Km11+00 - Km15+270", ChieuDaiKm = 4.30m, LoaiDuong = 4, GhiChu = "Hiện trạng" },
        new() { TuyenDuong = "QL.14E", DiaPhan = "Hiệp Đức - Phước Sơn", TuKmDenKm = "Km15+270 - Km89+432", ChieuDaiKm = 74.00m, LoaiDuong = 4, GhiChu = "Ngã ba Làng Hồi" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Hội An", TuKmDenKm = "Km0+00 - Km10+520", ChieuDaiKm = 10.50m, LoaiDuong = 3, GhiChu = "Cửa Đại" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Duy Xuyên", TuKmDenKm = "Km12+520 - Km17+500", ChieuDaiKm = 5.00m, LoaiDuong = 3, GhiChu = "Nam Phước" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Duy Xuyên", TuKmDenKm = "Km18+00 - Km24+910", ChieuDaiKm = 6.90m, LoaiDuong = 2, GhiChu = "Trùng QL.1" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Nông Sơn", TuKmDenKm = "Km24+910 - Km43+750", ChieuDaiKm = 18.80m, LoaiDuong = 3, GhiChu = "Kiểm Lâm" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Nông Sơn", TuKmDenKm = "Km43+750 - Km54+410", ChieuDaiKm = 10.70m, LoaiDuong = 4, GhiChu = "Trung Phước" },
        new() { TuyenDuong = "QL.14H", DiaPhan = "Quế Phước", TuKmDenKm = "Km60+220 - Km73+540", ChieuDaiKm = 13.70m, LoaiDuong = 5, GhiChu = "Quế Phước" },
        new() { TuyenDuong = "QL.40B", DiaPhan = "Tam Kỳ", TuKmDenKm = "Km1+770 - Km13+765", ChieuDaiKm = 12.00m, LoaiDuong = 3, GhiChu = "Tam Kỳ" },
        new() { TuyenDuong = "QL.40B", DiaPhan = "Phú Ninh", TuKmDenKm = "Km13+765 - Km32+300", ChieuDaiKm = 18.50m, LoaiDuong = 3, GhiChu = "Tiên Phước" },
        new() { TuyenDuong = "QL.40B", DiaPhan = "Bắc Trà My", TuKmDenKm = "Km32+300 - Km54+540", ChieuDaiKm = 22.20m, LoaiDuong = 5, GhiChu = "Trà My" },
        new() { TuyenDuong = "QL.40B", DiaPhan = "Nam Trà My", TuKmDenKm = "Km54+540 - Km125+00", ChieuDaiKm = 70.50m, LoaiDuong = 4, GhiChu = "Đoạn Trà Don" },
        new() { TuyenDuong = "QL.14G", DiaPhan = "Hòa Vang - Đông Giang", TuKmDenKm = "Km0+00 - Km16+646", ChieuDaiKm = 16.60m, LoaiDuong = 4, GhiChu = "Bà Nà" },
        new() { TuyenDuong = "QL.14G", DiaPhan = "Đông Giang", TuKmDenKm = "Km25+00 - Km66+00", ChieuDaiKm = 41.00m, LoaiDuong = 5, GhiChu = "Đông Giang" },
        new() { TuyenDuong = "Nam Hải Vân - Túy Loan", DiaPhan = "Liên Chiểu - Hòa Vang", TuKmDenKm = "Km12+182 - Km16+534", ChieuDaiKm = 4.35m, LoaiDuong = 3, GhiChu = "Túy Loan" },
        new() { TuyenDuong = "ĐT.601", DiaPhan = "Liên Chiểu - Hòa Vang", TuKmDenKm = "Km0+00 - Km35+681", ChieuDaiKm = 35.70m, LoaiDuong = 4, GhiChu = "Hòa Khánh - Hải Vân" },
        new() { TuyenDuong = "ĐT.602", DiaPhan = "Hòa Vang", TuKmDenKm = "Km0+00 - Km9+800", ChieuDaiKm = 9.80m, LoaiDuong = 3, GhiChu = "Âu Cơ - Bà Nà Suối Mơ" },
        new() { TuyenDuong = "ĐT.603", DiaPhan = "Điện Bàn", TuKmDenKm = "Km0+00 - Km4+270", ChieuDaiKm = 4.30m, LoaiDuong = 3, GhiChu = "Ngã ba Tứ Câu" },
        new() { TuyenDuong = "ĐT.603B", DiaPhan = "Điện Bàn - Hội An", TuKmDenKm = "Km0+00 - Km11+600", ChieuDaiKm = 11.60m, LoaiDuong = 2, GhiChu = "Trường Sa - Cửa Đại" },
        new() { TuyenDuong = "ĐT.605", DiaPhan = "Hòa Vang - Điện Bàn", TuKmDenKm = "Km0+00 - Km13+911", ChieuDaiKm = 13.90m, LoaiDuong = 3, GhiChu = "QL.1 - ĐT.609" },
        new() { TuyenDuong = "ĐT.607", DiaPhan = "Ngũ Hành Sơn - Hội An", TuKmDenKm = "Km0+00 - Km15+260", ChieuDaiKm = 15.30m, LoaiDuong = 2, GhiChu = "Trần Đại Nghĩa - Hội An" },
        new() { TuyenDuong = "ĐT.608", DiaPhan = "Điện Bàn - Hội An", TuKmDenKm = "Km0+00 - Km8+00", ChieuDaiKm = 8.00m, LoaiDuong = 3, GhiChu = "Vĩnh Điện - Thanh Hà" },
        new() { TuyenDuong = "ĐT.609", DiaPhan = "Điện Bàn - Đại Lộc", TuKmDenKm = "Km0+00 - Km74+300", ChieuDaiKm = 74.30m, LoaiDuong = 3, GhiChu = "Vĩnh Điện - Ái Nghĩa" },
        new() { TuyenDuong = "ĐT.619", DiaPhan = "Điện Bàn - Tam Kỳ - Núi Thành", TuKmDenKm = "Km0+00 - Km69+350", ChieuDaiKm = 69.40m, LoaiDuong = 1, GhiChu = "Đường Võ Chí Công (Ven biển)" }
    };

    // ==========================================
    // 4. THUẬT TOÁN TÍNH HAO PHÍ CỰ LY Ô TÔ
    // ==========================================
    public static (decimal caXe, decimal kmQuyDoiL1, decimal kmQuyDoiL2, decimal kmQuyDoiL3, decimal kmQuyDoiL4) TinhHaoPhiCaXeOTo(
        DinhMucVCOToItem dm, 
        List<CungDuongVanChuyen> cungDuongs)
    {
        if (cungDuongs == null || cungDuongs.Count == 0) return (0, 0, 0, 0, 0);

        decimal q1 = 0; // Tích lũy trong phạm vi <= 1km
        decimal q2 = 0; // Tích lũy trong phạm vi 1km < L <= 10km
        decimal q3 = 0; // Tích lũy trong phạm vi 10km < L <= 60km
        decimal q4 = 0; // Tích lũy trong phạm vi > 60km

        decimal currentStart = 0;
        foreach (var cd in cungDuongs)
        {
            if (cd.CuLyKm <= 0) continue;
            decimal currentEnd = currentStart + cd.CuLyKm;
            decimal k = cd.HeSoK;

            // Nấc 1: [0, 1]
            decimal l1 = Math.Max(0, Math.Min(currentEnd, 1m) - Math.Max(currentStart, 0m));
            // Nấc 2: [1, 10]
            decimal l2 = Math.Max(0, Math.Min(currentEnd, 10m) - Math.Max(currentStart, 1m));
            // Nấc 3: [10, 60]
            decimal l3 = Math.Max(0, Math.Min(currentEnd, 60m) - Math.Max(currentStart, 10m));
            // Nấc 4: [60, +oo)
            decimal l4 = Math.Max(0, currentEnd - Math.Max(currentStart, 60m));

            q1 += l1 * k;
            q2 += l2 * k;
            q3 += l3 * k;
            q4 += l4 * k;

            currentStart = currentEnd;
        }

        decimal caXe = dm.Dm1 * q1 + dm.Dm2 * q2 + dm.Dm3 * q3 + (dm.Dm3 * 0.95m) * q4;
        return (caXe, q1, q2, q3, q4);
    }

    // ==========================================
    // 5. THUẬT TOÁN TÍNH HAO PHÍ VẬN CHUYỂN BỘ
    // ==========================================
    public static decimal TinhHaoPhiNhanCongBo(
        DinhMucVCBoItem dm, 
        decimal cuLyMet, 
        decimal heSoDiaHinh = 1.0m, 
        int soTang = 1)
    {
        if (cuLyMet <= 0) return 0;
        
        // Cự ly khởi điểm 10m
        decimal haoPhi = dm.Dm10m;
        if (cuLyMet > 10m)
        {
            // Mỗi 10m tiếp theo
            decimal cuLyConLai = Math.Min(cuLyMet, 300m) - 10m;
            decimal soDoan10m = cuLyConLai / 10m;
            haoPhi += soDoan10m * dm.DmTiepTheo;
        }

        // Nhân hệ số địa hình
        haoPhi *= (heSoDiaHinh > 0 ? heSoDiaHinh : 1.0m);

        // Vận chuyển lên tầng cao (từ tầng 2 trở lên mỗi tầng * 1.1)
        if (soTang > 1)
        {
            decimal hsTang = (decimal)Math.Pow(1.1, soTang - 1);
            haoPhi *= hsTang;
        }

        return haoPhi;
    }

    // ==========================================
    // 6. NHẬN DIỆN VẬT LIỆU CHO Ô TÔ VÀ BỘ
    // ==========================================
    public static DinhMucVCOToItem? NhanDienOTo(string tenVatLieu)
    {
        if (string.IsNullOrWhiteSpace(tenVatLieu)) return null;
        var t = tenVatLieu.Trim().ToLower();

        // 1. Cát
        if (t.Contains("cát") || t.Contains("cat"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2311"); // 7t default

        // 2. Đá dăm / đá 1x2 / đá 2x4 / sỏi
        if (t.Contains("đá 1x2") || t.Contains("đá 2x4") || t.Contains("đá 4x6") || t.Contains("đá dăm") || t.Contains("sỏi"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2341"); // 7t

        // 3. Đá hộc
        if (t.Contains("đá hộc"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2351");

        // 4. Đất
        if (t.Contains("đất") || t.Contains("dat"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2321");

        // 5. Xi măng
        if (t.Contains("xi măng") || t.Contains("xi mang") || t.Contains("pcb"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2441");

        // 6. Thép
        if (t.Contains("thép") || t.Contains("thep") || t.Contains("sắt") || t.Contains("dây thép") || t.Contains("que hàn"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2451");

        // 7. Gạch xây
        if (t.Contains("gạch") && !t.Contains("lát") && !t.Contains("ốp"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2411");

        // 8. Gạch ốp lát
        if (t.Contains("gạch") && (t.Contains("lát") || t.Contains("ốp") || t.Contains("ceramic")))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2421");

        // 9. Ngói
        if (t.Contains("ngói") || t.Contains("ngoi"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2431");

        // 10. Ống cống
        if (t.Contains("cống") || t.Contains("cong"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2611");

        // 11. Cọc, cột bê tông
        if (t.Contains("cọc") || t.Contains("cột") || t.Contains("cot") || t.Contains("coc"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2711");

        // 12. Cấu kiện bê tông
        if (t.Contains("cấu kiện") || t.Contains("bê tông") || t.Contains("tam đan") || t.Contains("gối cống"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2511");

        // 13. Nhựa đường
        if (t.Contains("nhựa đường") || t.Contains("bitum"))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2461");

        // 14. Gỗ
        if (t.Contains("gỗ") || t.Contains("go "))
            return DanhSachOTo.FirstOrDefault(x => x.MaHieu == "AM.2471");

        return null;
    }

    public static DinhMucVCBoItem? NhanDienBo(string tenVatLieu)
    {
        if (string.IsNullOrWhiteSpace(tenVatLieu)) return null;
        var t = tenVatLieu.Trim().ToLower();

        if (t.Contains("cát") || t.Contains("cat")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2101");
        if (t.Contains("gạch") && (t.Contains("lát") || t.Contains("ốp"))) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2106");
        if (t.Contains("gạch")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2105");
        if (t.Contains("ngói")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2107");
        if (t.Contains("đất") || t.Contains("dat")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2102");
        if (t.Contains("xi măng") || t.Contains("pcb")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2108");
        if (t.Contains("thép") || t.Contains("sắt") || t.Contains("que hàn") || t.Contains("dây thép")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2112");
        if (t.Contains("cọc") || t.Contains("cừ")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2110");
        if (t.Contains("tre") || t.Contains("cây chống")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2111");
        if (t.Contains("gỗ")) return DanhSachBo.FirstOrDefault(x => x.MaHieu == "AM.2109");

        return null;
    }

    // ==========================================
    // 7. QUY ĐỔI ĐƠN VỊ TÍNH VẬN CHUYỂN
    // ==========================================
    public static decimal TinhHeSoQuyDoiOTo(string donViVatTu, string donViDinhMuc)
    {
        // Định mức tính cho 10m3 hoặc 10tấn
        // Quy đổi về 1 đơn vị vật tư:
        // Đơn vị định mức thường là "10m³" hoặc "10tấn"
        var dvVT = (donViVatTu ?? "").Trim().ToLower();
        var dvDM = (donViDinhMuc ?? "").Trim().ToLower();

        if (dvDM.Contains("10m3") || dvDM.Contains("10m³"))
        {
            if (dvVT == "m3" || dvVT == "m³") return 0.1m; // 1 m3 = 0.1 x 10m3
            if (dvVT == "lít" || dvVT == "lit") return 0.0001m; // 1 lít = 0.001 m3 = 0.0001 x 10m3
        }

        if (dvDM.Contains("10tấn") || dvDM.Contains("10tan"))
        {
            if (dvVT == "tấn" || dvVT == "tan" || dvVT == "t") return 0.1m; // 1 tấn = 0.1 x 10tấn
            if (dvVT == "kg") return 0.0001m; // 1 kg = 0.001 tấn = 0.0001 x 10tấn
            if (dvVT == "tạ") return 0.01m;
            if (dvVT == "bao") return 0.005m; // 1 bao 50kg = 0.05 tấn = 0.005 x 10tấn
            if (dvVT == "viên" || dvVT == "v") return 0.000015m; // Ước tính 1 viên gạch 1.5kg
            if (dvVT == "1000v" || dvVT == "nghìn viên") return 0.015m;
            if (dvVT == "đoạn" || dvVT == "cấu kiện" || dvVT == "ống" || dvVT == "cột" || dvVT == "cọc") return 0.1m; // Mặc định 1 cấu kiện = 1 tấn nếu tính theo tấn
        }

        return 0.1m; // Mặc định 1 đơn vị = 1/10 của 10 đơn vị
    }

    public static decimal TinhHeSoQuyDoiBo(string donViVatTu, string donViDinhMuc)
    {
        return DinhMucBocXepDatabase.TinhHeSoQuyDoi(donViVatTu, donViDinhMuc);
    }
}
