using System;
using System.Collections.Generic;

namespace AIE.Core.Models
{
    /// <summary>
    /// Cơ sở dữ liệu định mức chi phí Quản lý dự án và Tư vấn đầu tư xây dựng
    /// theo Thông tư số 38/2026/TT-BXD (Phụ lục VIII)
    /// </summary>
    public static class DinhMucTT38Database
    {
        // =========================================================================
        // BẢNG 1.1: ĐỊNH MỨC CHI PHÍ QUẢN LÝ DỰ ÁN (Phần trăm %)
        // Căn cứ tính: Chi phí xây dựng và thiết bị trước thuế (G_XD + G_TB)
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, >=30000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoQLDA = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m, 20000m, 30000m
        };

        public static readonly Dictionary<string, decimal[]> TiLeQLDA = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 3.283m, 2.766m, 2.213m, 1.839m, 1.492m, 1.157m, 0.902m, 0.702m, 0.504m, 0.380m, 0.285m, 0.228m },
            ["Công nghiệp"] = new decimal[] { 3.450m, 2.907m, 2.326m, 1.933m, 1.568m, 1.216m, 0.948m, 0.738m, 0.530m, 0.400m, 0.300m, 0.240m },
            ["Giao thông"] = new decimal[] { 2.951m, 2.486m, 1.989m, 1.653m, 1.341m, 1.040m, 0.811m, 0.631m, 0.453m, 0.342m, 0.256m, 0.205m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 3.117m, 2.626m, 2.101m, 1.746m, 1.417m, 1.099m, 0.857m, 0.667m, 0.479m, 0.361m, 0.271m, 0.217m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 2.787m, 2.348m, 1.879m, 1.561m, 1.267m, 0.982m, 0.766m, 0.596m, 0.428m, 0.323m, 0.242m, 0.194m }
        };

        // =========================================================================
        // BẢNG 2.4: ĐỊNH MỨC LẬP BÁO CÁO KINH TẾ - KỸ THUẬT (Quy mô < 15 tỷ)
        // Mốc quy mô: <=1, 2, 5, 10, >=15 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoBaoCaoKTKT = new decimal[] { 1m, 2m, 5m, 10m, 15m };

        public static readonly Dictionary<string, decimal[]> TiLeBaoCaoKTKT = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 6.50m, 5.80m, 4.80m, 3.90m, 3.20m },
            ["Công nghiệp"] = new decimal[] { 6.80m, 6.10m, 5.10m, 4.20m, 3.50m },
            ["Giao thông"] = new decimal[] { 5.20m, 4.60m, 3.80m, 3.10m, 2.50m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 5.80m, 5.10m, 4.20m, 3.40m, 2.80m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 5.50m, 4.80m, 4.00m, 3.20m, 2.60m }
        };

        // =========================================================================
        // BẢNG 2.2: ĐỊNH MỨC LẬP BÁO CÁO NGHIÊN CỨU KHẢ THI (FS)
        // Căn cứ: G_XD + G_TB
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoFS = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m
        };

        public static readonly Dictionary<string, decimal[]> TiLeLapFS = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 1.050m, 0.880m, 0.700m, 0.580m, 0.470m, 0.360m, 0.280m, 0.220m, 0.160m, 0.120m },
            ["Công nghiệp"] = new decimal[] { 1.150m, 0.960m, 0.760m, 0.630m, 0.510m, 0.390m, 0.300m, 0.240m, 0.170m, 0.130m },
            ["Giao thông"] = new decimal[] { 0.850m, 0.710m, 0.560m, 0.470m, 0.380m, 0.290m, 0.230m, 0.180m, 0.130m, 0.100m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 0.950m, 0.800m, 0.630m, 0.520m, 0.420m, 0.320m, 0.250m, 0.200m, 0.140m, 0.110m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.900m, 0.750m, 0.600m, 0.490m, 0.400m, 0.310m, 0.240m, 0.190m, 0.140m, 0.100m }
        };

        // =========================================================================
        // BẢNG 2.7 - 2.16: ĐỊNH MỨC THIẾT KẾ XÂY DỰNG CÔNG TRÌNH (Thiết kế 2 bước - TKBVTC)
        // Căn cứ: Chi phí xây dựng trước thuế (G_XD)
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoThietKe = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m
        };

        // Bảng tra Thiết kế BVTC theo: Loại CT + Cấp CT (Cấp ĐB, Cấp I, Cấp II, Cấp III, Cấp IV)
        // Mẫu hệ số tỷ lệ cho Cấp III (cấp phổ biến nhất), các cấp khác nhân tỷ lệ cấp tương ứng:
        // Cấp ĐB: x1.45; Cấp I: x1.30; Cấp II: x1.15; Cấp III: x1.00; Cấp IV: x0.85
        public static readonly Dictionary<string, decimal[]> TiLeThietKeCap3 = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 3.20m, 2.70m, 2.15m, 1.80m, 1.45m, 1.12m, 0.88m, 0.68m, 0.48m, 0.36m },
            ["Công nghiệp"] = new decimal[] { 3.40m, 2.85m, 2.28m, 1.90m, 1.53m, 1.18m, 0.92m, 0.72m, 0.51m, 0.38m },
            ["Giao thông"] = new decimal[] { 2.45m, 2.05m, 1.65m, 1.38m, 1.12m, 0.86m, 0.68m, 0.53m, 0.38m, 0.28m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 2.80m, 2.35m, 1.88m, 1.57m, 1.26m, 0.98m, 0.76m, 0.60m, 0.42m, 0.32m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 2.65m, 2.22m, 1.77m, 1.48m, 1.19m, 0.92m, 0.72m, 0.56m, 0.40m, 0.30m }
        };

        // =========================================================================
        // BẢNG 2.19: ĐỊNH MỨC THẨM TRA THIẾT KẾ XÂY DỰNG
        // Căn cứ: G_XD
        // Mốc quy mô: <=10, 20, 50, 100, 200, 500, 1000, 2000, 5000, >=10000 tỷ
        // =========================================================================
        public static readonly decimal[] MocQuyMoThamTraTK = new decimal[]
        {
            10m, 20m, 50m, 100m, 200m, 500m, 1000m, 2000m, 5000m, 10000m
        };

        public static readonly Dictionary<string, decimal[]> TiLeThamTraTK = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.210m, 0.177m, 0.141m, 0.117m, 0.095m, 0.073m, 0.057m, 0.045m, 0.032m, 0.024m },
            ["Công nghiệp"] = new decimal[] { 0.230m, 0.194m, 0.155m, 0.129m, 0.104m, 0.081m, 0.063m, 0.049m, 0.035m, 0.026m },
            ["Giao thông"] = new decimal[] { 0.160m, 0.135m, 0.108m, 0.090m, 0.073m, 0.056m, 0.044m, 0.034m, 0.025m, 0.018m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 0.180m, 0.152m, 0.121m, 0.101m, 0.081m, 0.063m, 0.049m, 0.039m, 0.028m, 0.021m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.175m, 0.148m, 0.118m, 0.098m, 0.079m, 0.061m, 0.048m, 0.037m, 0.027m, 0.020m }
        };

        // =========================================================================
        // BẢNG 2.20: ĐỊNH MỨC THẨM TRA DỰ TOÁN XÂY DỰNG
        // Căn cứ: G_XD
        // =========================================================================
        public static readonly Dictionary<string, decimal[]> TiLeThamTraDuToan = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 0.185m, 0.156m, 0.124m, 0.103m, 0.083m, 0.064m, 0.050m, 0.039m, 0.028m, 0.021m },
            ["Công nghiệp"] = new decimal[] { 0.205m, 0.173m, 0.138m, 0.115m, 0.093m, 0.072m, 0.056m, 0.044m, 0.031m, 0.023m },
            ["Giao thông"] = new decimal[] { 0.145m, 0.122m, 0.097m, 0.081m, 0.066m, 0.051m, 0.040m, 0.031m, 0.022m, 0.016m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 0.160m, 0.135m, 0.108m, 0.090m, 0.072m, 0.056m, 0.044m, 0.034m, 0.025m, 0.018m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 0.155m, 0.131m, 0.104m, 0.087m, 0.070m, 0.054m, 0.042m, 0.033m, 0.024m, 0.018m }
        };

        // =========================================================================
        // BẢNG 2.24: ĐỊNH MỨC GIÁM SÁT THI CÔNG XÂY DỰNG
        // Căn cứ: G_XD
        // =========================================================================
        public static readonly Dictionary<string, decimal[]> TiLeGiamSatThiCong = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dân dụng"] = new decimal[] { 2.85m, 2.40m, 1.92m, 1.60m, 1.30m, 1.01m, 0.78m, 0.61m, 0.44m, 0.33m },
            ["Công nghiệp"] = new decimal[] { 3.05m, 2.57m, 2.06m, 1.71m, 1.39m, 1.08m, 0.84m, 0.65m, 0.47m, 0.35m },
            ["Giao thông"] = new decimal[] { 2.45m, 2.07m, 1.65m, 1.37m, 1.11m, 0.86m, 0.67m, 0.52m, 0.38m, 0.28m },
            ["Nông nghiệp & PTNT"] = new decimal[] { 2.65m, 2.23m, 1.79m, 1.48m, 1.20m, 0.93m, 0.73m, 0.57m, 0.41m, 0.30m },
            ["Hạ tầng kỹ thuật"] = new decimal[] { 2.55m, 2.15m, 1.72m, 1.43m, 1.16m, 0.90m, 0.70m, 0.55m, 0.39m, 0.29m }
        };

        // =========================================================================
        // BẢNG 2.21 - 2.23: LẬP HSMT VÀ ĐÁNH GIÁ HSDT THI CÔNG
        // Căn cứ: G_XD (Mặc định: Lập HSMT = 0.35 * QLDA; Đánh giá = 0.40 * QLDA)
        // =========================================================================
        public static readonly decimal[] MocQuyMoHSMT = new decimal[] { 10m, 20m, 50m, 100m, 200m, 500m, 1000m };
        public static readonly decimal[] TiLeLapHSMT = new decimal[] { 0.38m, 0.32m, 0.26m, 0.21m, 0.17m, 0.13m, 0.10m };
        public static readonly decimal[] TiLeDanhGiaHSDT = new decimal[] { 0.42m, 0.36m, 0.29m, 0.24m, 0.19m, 0.15m, 0.11m };

        /// <summary>
        /// Hệ số điều chỉnh cấp công trình cho Thiết kế, Giám sát, Thẩm tra
        /// </summary>
        public static decimal GetHeSoCapCongTrinh(string capCT)
        {
            if (string.IsNullOrEmpty(capCT)) return 1.0m;
            if (capCT.Contains("đặc biệt") || capCT.Contains("ĐB") || capCT.Contains("Đặc biệt")) return 1.45m;
            if (capCT.Contains("I") && !capCT.Contains("II") && !capCT.Contains("III") && !capCT.Contains("IV")) return 1.25m;
            if (capCT.Contains("II") && !capCT.Contains("III")) return 1.12m;
            if (capCT.Contains("III")) return 1.00m; // Cấp III là chuẩn mốc
            if (capCT.Contains("IV")) return 0.85m;
            return 1.0m;
        }
    }
}
