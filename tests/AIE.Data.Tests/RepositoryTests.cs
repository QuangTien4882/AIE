using Xunit;
using AIE.Data;
using AIE.Data.Repositories;
using AIE.Core.Models;
using System.IO;
using System;

namespace AIE.Data.Tests;

public class RepositoryTests : IDisposable
{
    private readonly string _dbPath;
    private readonly AieDbContext _context;

    public RepositoryTests()
    {
        // Use a temporary DB for tests
        _dbPath = Path.Combine(Path.GetTempPath(), $"AIE_Test_{Guid.NewGuid()}.sqlite");
        _context = new AieDbContext(_dbPath);
        _context.InitializeDatabase();
    }

    [Fact]
    public void Can_Insert_And_Retrieve_VatLieu()
    {
        var repo = new VatLieuRepository(_context);
        var vl = new VatLieu
        {
            MaVL = "V08770",
            TenVL = "Xi măng PCB40",
            DonVi = "kg",
            DonGia = 1741,
            NhaSanXuat = "Hải Vân",
            GhiChu = "Giá tháng 7/2026",
            NgayCapNhat = DateTime.Now
        };

        repo.Upsert(vl);
        
        var retrieved = repo.GetByMa("V08770");
        Assert.NotNull(retrieved);
        Assert.Equal("Xi măng PCB40", retrieved.TenVL);
        Assert.Equal(1741, retrieved.DonGia);

        // Test Upsert update
        vl.DonGia = 1800;
        repo.Upsert(vl);

        var updated = repo.GetByMa("V08770");
        Assert.Equal(1800, updated?.DonGia);
    }

    [Fact]
    public void Can_Insert_And_Retrieve_CongTacXayDung()
    {
        var repo = new CongTacRepository(_context);
        var ct = new CongTacXayDung
        {
            MaHieu = "AF.11112",
            TenCongTac = "Bê tông lót móng",
            DonVi = "m3",
            DanhSachHaoPhi = new List<HaoPhi>
            {
                new HaoPhi { LoaiHaoPhi = Core.Enums.LoaiHaoPhi.VL, MaHieuHP = "V08770", TenHaoPhi = "Xi măng PCB40", DonVi = "kg", DinhMuc = 234.725m }
            }
        };

        repo.Insert(ct);

        var retrieved = repo.GetByMaHieu("AF.11112");
        Assert.NotNull(retrieved);
        Assert.Equal("Bê tông lót móng", retrieved.TenCongTac);
        Assert.Single(retrieved.DanhSachHaoPhi);
        Assert.Equal("V08770", retrieved.DanhSachHaoPhi[0].MaHieuHP);
        Assert.Equal(234.725m, retrieved.DanhSachHaoPhi[0].DinhMuc);
    }

    public void Dispose()
    {
        if (File.Exists(_dbPath))
        {
            try { File.Delete(_dbPath); } catch { }
        }
    }
}
