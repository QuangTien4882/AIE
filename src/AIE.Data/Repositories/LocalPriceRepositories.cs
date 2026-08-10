using AIE.Core.Models;
using Dapper;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Data.Repositories;

/// <summary>
/// Repository quản lý giá Vật liệu địa phương (Đà Nẵng).
/// </summary>
public class VatLieuRepository
{
    private readonly AieDbContext _context;

    public VatLieuRepository(AieDbContext context)
    {
        _context = context;
    }

    public IEnumerable<VatLieu> GetAll()
    {
        using var connection = _context.GetConnection();
        return connection.Query<VatLieu>("SELECT * FROM VatLieu");
    }

    public VatLieu? GetByMa(string maVL)
    {
        using var connection = _context.GetConnection();
        return connection.QueryFirstOrDefault<VatLieu>(
            "SELECT * FROM VatLieu WHERE MaVL = @MaVL", new { MaVL = maVL });
    }

    public void Upsert(VatLieu vl)
    {
        using var connection = _context.GetConnection();
        var sql = @"
            INSERT OR REPLACE INTO VatLieu (MaVL, TenVL, DonVi, DonGia, NhaSanXuat, GhiChu, NgayCapNhat)
            VALUES (@MaVL, @TenVL, @DonVi, @DonGia, @NhaSanXuat, @GhiChu, @NgayCapNhat);
        ";
        connection.Execute(sql, vl);
    }
}

/// <summary>
/// Repository quản lý giá Nhân công địa phương.
/// </summary>
public class NhanCongRepository
{
    private readonly AieDbContext _context;

    public NhanCongRepository(AieDbContext context)
    {
        _context = context;
    }

    public IEnumerable<NhanCong> GetAll()
    {
        using var connection = _context.GetConnection();
        return connection.Query<NhanCong>("SELECT * FROM NhanCong");
    }

    public NhanCong? GetByMa(string maNC)
    {
        using var connection = _context.GetConnection();
        return connection.QueryFirstOrDefault<NhanCong>(
            "SELECT * FROM NhanCong WHERE MaNC = @MaNC", new { MaNC = maNC });
    }

    public void Upsert(NhanCong nc)
    {
        using var connection = _context.GetConnection();
        var sql = @"
            INSERT OR REPLACE INTO NhanCong (MaNC, TenNC, Nhom, DonVi, DonGia, GhiChu, NgayCapNhat)
            VALUES (@MaNC, @TenNC, @Nhom, @DonVi, @DonGia, @GhiChu, @NgayCapNhat);
        ";
        connection.Execute(sql, nc);
    }
}

/// <summary>
/// Repository quản lý giá Máy thi công địa phương.
/// </summary>
public class MayThiCongRepository
{
    private readonly AieDbContext _context;

    public MayThiCongRepository(AieDbContext context)
    {
        _context = context;
    }

    public IEnumerable<MayThiCong> GetAll()
    {
        using var connection = _context.GetConnection();
        return connection.Query<MayThiCong>("SELECT * FROM MayThiCong");
    }

    public MayThiCong? GetByMa(string maMay)
    {
        using var connection = _context.GetConnection();
        return connection.QueryFirstOrDefault<MayThiCong>(
            "SELECT * FROM MayThiCong WHERE MaMay = @MaMay", new { MaMay = maMay });
    }

    public void Upsert(MayThiCong mtc)
    {
        using var connection = _context.GetConnection();
        var sql = @"
            INSERT OR REPLACE INTO MayThiCong (MaMay, TenMay, DonVi, DonGia, GhiChu, NgayCapNhat)
            VALUES (@MaMay, @TenMay, @DonVi, @DonGia, @GhiChu, @NgayCapNhat);
        ";
        connection.Execute(sql, mtc);
    }
}
