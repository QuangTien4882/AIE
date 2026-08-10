using AIE.Core.Models;

namespace AIE.Core.Services.Shared;

/// <summary>
/// Nội suy tuyến tính cho định mức % QLDA / Tư vấn.
/// Khi quy mô nằm giữa 2 mốc, nội suy tuyến tính để tính tỉ lệ % chính xác.
/// </summary>
public static class InterpolationHelper
{
    /// <summary>
    /// Tìm tỉ lệ % QLDA theo quy mô dự án.
    /// Nội suy tuyến tính nếu quy mô nằm giữa 2 mốc.
    /// </summary>
    /// <param name="danhSachDinhMuc">Danh sách định mức % theo mốc quy mô (đã sắp xếp tăng)</param>
    /// <param name="quyMo">Quy mô dự án (tỷ đồng)</param>
    /// <returns>Tỉ lệ % sau nội suy</returns>
    public static decimal NoiSuyTiLe(IReadOnlyList<DinhMucQLDA> danhSachDinhMuc, decimal quyMo)
    {
        if (danhSachDinhMuc.Count == 0)
            throw new ArgumentException("Danh sách định mức rỗng.");

        // Sắp xếp theo QuyMoMin tăng dần
        var sorted = danhSachDinhMuc.OrderBy(x => x.QuyMoMin).ToList();

        // Nếu quy mô nhỏ hơn mốc đầu tiên → lấy tỉ lệ đầu
        if (quyMo <= sorted[0].QuyMoMin)
            return sorted[0].TiLe;

        // Nếu quy mô lớn hơn mốc cuối → lấy tỉ lệ cuối
        var last = sorted[sorted.Count - 1];
        if (last.QuyMoMax.HasValue && quyMo >= last.QuyMoMax.Value)
            return last.TiLe;

        // Tìm 2 mốc kẹp
        for (int i = 0; i < sorted.Count; i++)
        {
            var dm = sorted[i];
            decimal min = dm.QuyMoMin;
            decimal max = dm.QuyMoMax ?? decimal.MaxValue;

            if (quyMo >= min && quyMo <= max)
            {
                // Nằm đúng trong 1 khoảng → lấy tỉ lệ đó
                // Nhưng nếu có mốc kế tiếp, nội suy giữa 2 mốc
                if (i + 1 < sorted.Count && quyMo > min)
                {
                    var next = sorted[i + 1];
                    return NoiSuyTuyenTinh(min, dm.TiLe, next.QuyMoMin, next.TiLe, quyMo);
                }
                return dm.TiLe;
            }
        }

        // Fallback: lấy mốc cuối
        return sorted[sorted.Count - 1].TiLe;
    }

    /// <summary>
    /// Tìm tỉ lệ % tư vấn theo quy mô.
    /// </summary>
    public static decimal NoiSuyTiLeTuVan(IReadOnlyList<DinhMucTuVan> danhSachDinhMuc, decimal quyMo)
    {
        if (danhSachDinhMuc.Count == 0)
            throw new ArgumentException("Danh sách định mức rỗng.");

        var sorted = danhSachDinhMuc.OrderBy(x => x.QuyMoMin).ToList();

        if (quyMo <= sorted[0].QuyMoMin)
            return sorted[0].TiLe;

        var last = sorted[sorted.Count - 1];
        if (last.QuyMoMax.HasValue && quyMo >= last.QuyMoMax.Value)
            return last.TiLe;

        for (int i = 0; i < sorted.Count; i++)
        {
            var dm = sorted[i];
            decimal min = dm.QuyMoMin;
            decimal max = dm.QuyMoMax ?? decimal.MaxValue;

            if (quyMo >= min && quyMo <= max)
            {
                if (i + 1 < sorted.Count && quyMo > min)
                {
                    var next = sorted[i + 1];
                    return NoiSuyTuyenTinh(min, dm.TiLe, next.QuyMoMin, next.TiLe, quyMo);
                }
                return dm.TiLe;
            }
        }

        return sorted[sorted.Count - 1].TiLe;
    }

    /// <summary>
    /// Nội suy tuyến tính giữa 2 điểm.
    /// </summary>
    public static decimal NoiSuyTuyenTinh(decimal x1, decimal y1, decimal x2, decimal y2, decimal x)
    {
        if (x2 == x1) return y1;
        decimal result = y1 + (y2 - y1) * (x - x1) / (x2 - x1);
        return Math.Round(result, 3);
    }
}
