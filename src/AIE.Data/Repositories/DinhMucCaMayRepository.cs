using AIE.Core.Models;
using Dapper;

namespace AIE.Data.Repositories;

public class DinhMucCaMayRepository
{
    private readonly AieDbContext _context;

    public DinhMucCaMayRepository(AieDbContext context)
    {
        _context = context;
    }

    public DinhMucCaMay_TT37? GetByMaMay(string maMay)
    {
        using var connection = _context.GetConnection();
        return connection.QueryFirstOrDefault<DinhMucCaMay_TT37>(
            "SELECT * FROM DinhMucCaMay_TT37 WHERE MaMay = @MaMay", new { MaMay = maMay });
    }

    public void Upsert(DinhMucCaMay_TT37 dm)
    {
        using var connection = _context.GetConnection();
        var sql = @"
            INSERT OR REPLACE INTO DinhMucCaMay_TT37 
            (MaMay, NguyenGia, KhauHao, SuaChua, ChiPhiKhac, DinhMucXang, DinhMucDiezel, DinhMucDien, SoLuongNhanCong, NhomNhanCong, SoCaNam, NhanCongString, ThanhPhanNhanCong)
            VALUES 
            (@MaMay, @NguyenGia, @KhauHao, @SuaChua, @ChiPhiKhac, @DinhMucXang, @DinhMucDiezel, @DinhMucDien, @SoLuongNhanCong, @NhomNhanCong, @SoCaNam, @NhanCongString, @ThanhPhanNhanCong);
        ";
        connection.Execute(sql, dm);
    }
}
