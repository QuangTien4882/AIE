using AIE.Core.Models;

namespace AIE.Core.Services.LapDuToan;

/// <summary>
/// Tính toán chi phí xây dựng theo Bảng 3.8 Thông tư 36/2026/TT-BXD và QĐ đính chính 1538/QĐ-BXD.
///   GXDTT = T + GT + TL  (Chi phí xây dựng trước thuế)
///   GTGT  = GXDTT × Thuế suất
///   GXD   = GXDTT + GTGT  (Chi phí xây dựng sau thuế — không cộng LT)
///   LT    = GXDTT × Tỷ lệ nhà tạm × (1 + Thuế suất)
/// </summary>
public class ChiPhiXayDungCalc
{
    /// <summary>
    /// Tính chi phí xây dựng đầy đủ từ chi phí trực tiếp và các tỉ lệ %.
    /// </summary>
    /// <param name="tongVL">Tổng chi phí vật liệu</param>
    /// <param name="tongNC">Tổng chi phí nhân công</param>
    /// <param name="tongMay">Tổng chi phí máy</param>
    /// <param name="tiLeCPC">Tỉ lệ % chi phí chung (VD: 7.3)</param>
    /// <param name="tiLeTT">Tỉ lệ % CP không XĐ KL (VD: 2.5)</param>
    /// <param name="tiLeTNCTTT">Tỉ lệ % TNCTTT (VD: 5.5)</param>
    /// <param name="tiLeGTGT">Thuế suất GTGT (VD: 8)</param>
    /// <param name="tiLeNhaTam">Tỉ lệ % nhà tạm (VD: 1.1)</param>
    /// <param name="coSoTinhCPC">Cơ sở tính CPC: "T" hoặc "NC"</param>
    public ChiPhiXayDung Tinh(
        decimal tongVL,
        decimal tongNC,
        decimal tongMay,
        decimal tiLeCPC,
        decimal tiLeTT,
        decimal tiLeTNCTTT,
        decimal tiLeGTGT,
        decimal tiLeNhaTam,
        string coSoTinhCPC = "T",
        string giaiDoan = "Lập dự toán xây dựng",
        string loaiNhaTam = "Công trình xây dựng còn lại",
        bool laVungSauXa = false)
    {
        var result = new ChiPhiXayDung
        {
            VL = tongVL,
            NC = tongNC,
            M = tongMay,
            TiLeCPC = tiLeCPC,
            TiLeTT = tiLeTT,
            TiLeTNCTTT = tiLeTNCTTT,
            TiLeGTGT = tiLeGTGT,
            TiLeNhaTam = tiLeNhaTam,
            GiaiDoan = giaiDoan,
            LoaiCongTrinhNhaTam = loaiNhaTam,
            LaVungSauXa = laVungSauXa
        };

        // 1. Chi phí trực tiếp: T = VL + NC + M
        decimal T = result.T;

        // 2. Chi phí chung: CPC = T × % (hoặc NC × %)
        decimal coSoCPC = coSoTinhCPC.Equals("NC", StringComparison.OrdinalIgnoreCase) ? tongNC : T;
        result.CPC = Math.Round(coSoCPC * tiLeCPC / 100m, 0, MidpointRounding.AwayFromZero);

        // 3. CP không XĐ được KL: TT = T × %
        result.TT = Math.Round(T * tiLeTT / 100m, 0, MidpointRounding.AwayFromZero);

        // 4. GT = CPC + TT (qua property GT)
        decimal GT = result.GT;

        // 5. Thu nhập chịu thuế tính trước: TL = (T + GT) × %
        result.TL = Math.Round((T + GT) * tiLeTNCTTT / 100m, 0, MidpointRounding.AwayFromZero);

        // 6. Chi phí xây dựng trước thuế: GXDTT = T + GT + TL (qua property GXDTT / G)
        decimal gxdtt = result.GXDTT;

        // 7. Thuế GTGT: GTGT = GXDTT × %
        result.GTGT = Math.Round(gxdtt * tiLeGTGT / 100m, 0, MidpointRounding.AwayFromZero);

        // 8. Chi phí xây dựng sau thuế: GXD = GXDTT + GTGT (qua property GXD / Gxd)

        // 9. Chi phí nhà tạm để ở và điều hành thi công (Mục V Bảng 3.8):
        // Theo QĐ 1538/QĐ-BXD: LT = GXDTT × Tỷ lệ × (1 + Thuế suất GTGT)
        decimal vatFactor = 1m + (tiLeGTGT / 100m);
        result.LT = Math.Round(gxdtt * (tiLeNhaTam / 100m) * vatFactor, 0, MidpointRounding.AwayFromZero);

        return result;
    }
}

