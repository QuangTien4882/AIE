using Dapper;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Data.Repositories;

/// <summary>
/// Repository quản lý Bộ đơn giá (BoDonGia) và giá theo từng bộ.
/// </summary>
public class BoDonGiaRepository
{
    private readonly AieDbContext _context;

    public BoDonGiaRepository(AieDbContext context)
    {
        _context = context;
    }

    // ==================== BoDonGia ====================

    public class BoDonGiaInfo
    {
        public int Id { get; set; }
        public string TenBo { get; set; } = string.Empty;
        public decimal GiaXang { get; set; }
        public decimal GiaDiezel { get; set; }
        public decimal GiaDien { get; set; }
        public string GhiChu { get; set; } = string.Empty;
        public string NgayTao { get; set; } = string.Empty;
    }

    public IEnumerable<BoDonGiaInfo> GetAll()
    {
        using var conn = _context.GetConnection();
        return conn.Query<BoDonGiaInfo>("SELECT * FROM BoDonGia ORDER BY Id DESC").ToList();
    }

    public BoDonGiaInfo GetById(int id)
    {
        using var conn = _context.GetConnection();
        return conn.QueryFirstOrDefault<BoDonGiaInfo>("SELECT * FROM BoDonGia WHERE Id = @Id", new { Id = id });
    }

    public int Create(string tenBo, decimal giaXang = 0, decimal giaDiezel = 0, decimal giaDien = 0, string ghiChu = "")
    {
        using var conn = _context.GetConnection();
        var sql = @"INSERT INTO BoDonGia (TenBo, GiaXang, GiaDiezel, GiaDien, GhiChu, NgayTao)
                    VALUES (@TenBo, @GiaXang, @GiaDiezel, @GiaDien, @GhiChu, @NgayTao);
                    SELECT last_insert_rowid();";
        return (int)conn.ExecuteScalar<long>(sql, new
        {
            TenBo = tenBo,
            GiaXang = giaXang,
            GiaDiezel = giaDiezel,
            GiaDien = giaDien,
            GhiChu = ghiChu,
            NgayTao = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        });
    }

    public void UpdateName(int boDonGiaId, string newName)
    {
        using var conn = _context.GetConnection();
        conn.Execute("UPDATE BoDonGia SET TenBo = @TenBo WHERE Id = @Id", new { TenBo = newName, Id = boDonGiaId });
    }

    public void UpdateFuelPrices(int boDonGiaId, decimal giaXang, decimal giaDiezel, decimal giaDien)
    {
        using var conn = _context.GetConnection();
        conn.Execute(
            "UPDATE BoDonGia SET GiaXang = @Xang, GiaDiezel = @Diezel, GiaDien = @Dien WHERE Id = @Id",
            new { Xang = giaXang, Diezel = giaDiezel, Dien = giaDien, Id = boDonGiaId });
    }

    public void Delete(int boDonGiaId)
    {
        using var conn = _context.GetConnection();
        conn.Execute("DELETE FROM GiaVatLieuTheoBo WHERE BoDonGiaId = @Id", new { Id = boDonGiaId });
        conn.Execute("DELETE FROM GiaNhanCongTheoBo WHERE BoDonGiaId = @Id", new { Id = boDonGiaId });
        conn.Execute("DELETE FROM GiaMayTheoBo WHERE BoDonGiaId = @Id", new { Id = boDonGiaId });
        conn.Execute("DELETE FROM BoDonGia WHERE Id = @Id", new { Id = boDonGiaId });
    }

    // ==================== Giá VL theo Bộ ====================

    public class GiaVatLieuBo
    {
        public string MaVL { get; set; } = string.Empty;
        public decimal GiaGoc { get; set; }
        public decimal CuocVC { get; set; }
        public decimal GiaHienTruong { get; set; }
    }

    public void SaveGiaVL(int boDonGiaId, string maVL, decimal giaGoc, decimal cuocVC, decimal giaHienTruong)
    {
        using var conn = _context.GetConnection();
        var exists = conn.ExecuteScalar<int>("SELECT COUNT(*) FROM GiaVatLieuTheoBo WHERE BoDonGiaId = @BoDonGiaId AND MaVL = @MaVL", new { BoDonGiaId = boDonGiaId, MaVL = maVL });
        if (exists > 0)
        {
            conn.Execute("UPDATE GiaVatLieuTheoBo SET GiaGoc = @GiaGoc, CuocVC = @CuocVC, GiaHienTruong = @GiaHienTruong WHERE BoDonGiaId = @BoDonGiaId AND MaVL = @MaVL", 
                new { BoDonGiaId = boDonGiaId, MaVL = maVL, GiaGoc = giaGoc, CuocVC = cuocVC, GiaHienTruong = giaHienTruong });
        }
        else
        {
            var sql = @"INSERT INTO GiaVatLieuTheoBo 
                        (BoDonGiaId, MaVL, GiaGoc, CuocVC, GiaHienTruong, DuocChon) 
                        VALUES (@BoDonGiaId, @MaVL, @GiaGoc, @CuocVC, @GiaHienTruong, 1)";
            conn.Execute(sql, new { BoDonGiaId = boDonGiaId, MaVL = maVL, GiaGoc = giaGoc, CuocVC = cuocVC, GiaHienTruong = giaHienTruong });
        }
    }

    public IEnumerable<GiaVatLieuBo> GetGiaVL(int boDonGiaId)
    {
        using var conn = _context.GetConnection();
        return conn.Query<GiaVatLieuBo>(
            "SELECT MaVL, GiaGoc, CuocVC, GiaHienTruong FROM GiaVatLieuTheoBo WHERE BoDonGiaId = @Id",
            new { Id = boDonGiaId }).ToList();
    }

    // ==================== Giá NC theo Bộ ====================

    public class GiaNhanCongBo
    {
        public string MaNC { get; set; } = string.Empty;
        public decimal DonGia { get; set; }
    }

    public void SaveGiaNC(int boDonGiaId, string maNC, decimal donGia)
    {
        using var conn = _context.GetConnection();
        conn.Execute(
            "INSERT OR REPLACE INTO GiaNhanCongTheoBo (BoDonGiaId, MaNC, DonGia) VALUES (@BoDonGiaId, @MaNC, @DonGia)",
            new { BoDonGiaId = boDonGiaId, MaNC = maNC, DonGia = donGia });
    }

    public IEnumerable<GiaNhanCongBo> GetGiaNC(int boDonGiaId)
    {
        using var conn = _context.GetConnection();
        return conn.Query<GiaNhanCongBo>(
            "SELECT MaNC, DonGia FROM GiaNhanCongTheoBo WHERE BoDonGiaId = @Id",
            new { Id = boDonGiaId }).ToList();
    }

    // ==================== Giá Máy theo Bộ ====================

    public class GiaMayBo
    {
        public string MaMay { get; set; } = string.Empty;
        public decimal DonGia { get; set; }
    }

    public void SaveGiaMay(int boDonGiaId, string maMay, decimal donGia)
    {
        using var conn = _context.GetConnection();
        conn.Execute(
            "INSERT OR REPLACE INTO GiaMayTheoBo (BoDonGiaId, MaMay, DonGia) VALUES (@BoDonGiaId, @MaMay, @DonGia)",
            new { BoDonGiaId = boDonGiaId, MaMay = maMay, DonGia = donGia });
    }

    public IEnumerable<GiaMayBo> GetGiaMay(int boDonGiaId)
    {
        using var conn = _context.GetConnection();
        return conn.Query<GiaMayBo>(
            "SELECT MaMay, DonGia FROM GiaMayTheoBo WHERE BoDonGiaId = @Id",
            new { Id = boDonGiaId }).ToList();
    }

    // ==================== Sao chép Bộ đơn giá ====================

    /// <summary>
    /// Sao chép toàn bộ giá từ bộ nguồn sang bộ đích.
    /// </summary>
    public void CopyPrices(int fromBoDonGiaId, int toBoDonGiaId)
    {
        using var conn = _context.GetConnection();

        // Copy VL
        var vlPrices = GetGiaVL(fromBoDonGiaId);
        foreach (var vl in vlPrices)
            SaveGiaVL(toBoDonGiaId, vl.MaVL, vl.GiaGoc, vl.CuocVC, vl.GiaHienTruong);

        // Copy NC
        var ncPrices = GetGiaNC(fromBoDonGiaId);
        foreach (var nc in ncPrices)
            SaveGiaNC(toBoDonGiaId, nc.MaNC, nc.DonGia);

        // Copy May
        var mayPrices = GetGiaMay(fromBoDonGiaId);
        foreach (var m in mayPrices)
            SaveGiaMay(toBoDonGiaId, m.MaMay, m.DonGia);
    }
}
