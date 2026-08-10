using AIE.Core.Models;

namespace AIE.Core.Services.LapDuToan;

/// <summary>
/// Tính toán chi phí xây dựng (GXD) theo công thức TT36/2026.
/// GXD = Gxd + LT
/// Gxd = G + GTGT
/// G   = T + GT + TL
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
        string coSoTinhCPC = "T")
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
            TiLeNhaTam = tiLeNhaTam
        };

        // T = VL + NC + M (tự tính qua property)
        decimal T = result.T;

        // Chi phí chung: CPC = T × % hoặc NC × %
        decimal coSoCPC = coSoTinhCPC.Equals("NC", StringComparison.OrdinalIgnoreCase) ? tongNC : T;
        result.CPC = Math.Round(coSoCPC * tiLeCPC / 100, 0, MidpointRounding.AwayFromZero);

        // CP không XĐ được KL: TT = T × %
        result.TT = Math.Round(T * tiLeTT / 100, 0, MidpointRounding.AwayFromZero);

        // GT = CPC + TT (tự tính qua property)
        decimal GT = result.GT;

        // TNCTTT: TL = (T + GT) × %
        result.TL = Math.Round((T + GT) * tiLeTNCTTT / 100, 0, MidpointRounding.AwayFromZero);

        // G = T + GT + TL (tự tính qua property)
        decimal G = result.G;

        // GTGT = G × thuế suất
        result.GTGT = Math.Round(G * tiLeGTGT / 100, 0, MidpointRounding.AwayFromZero);

        // Nhà tạm: LT = Gxd × %
        result.LT = Math.Round(result.Gxd * tiLeNhaTam / 100, 0, MidpointRounding.AwayFromZero);

        return result;
    }
}
