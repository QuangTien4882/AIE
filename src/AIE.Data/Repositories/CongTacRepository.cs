using AIE.Core.Models;
using Dapper;
using System.Collections.Generic;
using System.Linq;

namespace AIE.Data.Repositories;

/// <summary>
/// Repository quản lý Định mức Công tác xây dựng & Hao phí.
/// </summary>
public class CongTacRepository
{
    private readonly AieDbContext _context;

    public CongTacRepository(AieDbContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Lấy danh sách công tác xây dựng kèm theo chi tiết hao phí.
    /// </summary>
    public IEnumerable<CongTacXayDung> GetAllCongTac()
    {
        using var connection = _context.GetConnection();
        
        var sql = @"
            SELECT * FROM CongTacXayDung;
            SELECT * FROM HaoPhi;
        ";

        using var multi = connection.QueryMultiple(sql);
        var dsCongTac = multi.Read<CongTacXayDung>().ToList();
        var dsHaoPhi = multi.Read<HaoPhi>().ToList();

        // Gắn HaoPhi vào CongTac
        var lookup = dsHaoPhi.GroupBy(hp => hp.CongTacId)
                             .ToDictionary(g => g.Key, g => g.ToList());

        foreach (var ct in dsCongTac)
        {
            if (lookup.TryGetValue(ct.Id, out var haophis))
            {
                ct.DanhSachHaoPhi = haophis;
            }
        }

        return dsCongTac;
    }

    /// <summary>
    /// Đếm tổng số công tác trong DB.
    /// </summary>
    public int Count()
    {
        using var connection = _context.GetConnection();
        return connection.ExecuteScalar<int>("SELECT COUNT(*) FROM CongTacXayDung");
    }

    /// <summary>
    /// Lấy 1 công tác theo Mã hiệu (VD: AF.11112)
    /// </summary>
    public CongTacXayDung? GetByMaHieu(string maHieu)
    {
        using var connection = _context.GetConnection();
        
        var sqlCongTac = "SELECT * FROM CongTacXayDung WHERE MaHieu = @MaHieu;";
        var ct = connection.QueryFirstOrDefault<CongTacXayDung>(sqlCongTac, new { MaHieu = maHieu });

        if (ct != null)
        {
            var sqlHaoPhi = "SELECT * FROM HaoPhi WHERE CongTacId = @CongTacId;";
            ct.DanhSachHaoPhi = connection.Query<HaoPhi>(sqlHaoPhi, new { CongTacId = ct.Id }).ToList();
        }

        return ct;
    }

    /// <summary>
    /// Tìm kiếm công tác theo mã hiệu hoặc tên
    /// </summary>
    public IEnumerable<CongTacXayDung> SearchCongTac(string keyword)
    {
        using var connection = _context.GetConnection();
        var param = new { Keyword = $"%{keyword}%" };
        var sqlCongTac = "SELECT * FROM CongTacXayDung WHERE MaHieu LIKE @Keyword OR TenCongTac LIKE @Keyword LIMIT 100;";
        
        var dsCongTac = connection.Query<CongTacXayDung>(sqlCongTac, param).ToList();
        if (dsCongTac.Any())
        {
            var ids = dsCongTac.Select(x => x.Id).ToList();
            var sqlHaoPhi = "SELECT * FROM HaoPhi WHERE CongTacId IN @Ids;";
            var dsHaoPhi = connection.Query<HaoPhi>(sqlHaoPhi, new { Ids = ids }).ToList();

            var lookup = dsHaoPhi.GroupBy(hp => hp.CongTacId).ToDictionary(g => g.Key, g => g.ToList());
            foreach (var ct in dsCongTac)
            {
                if (lookup.TryGetValue(ct.Id, out var haophis))
                {
                    ct.DanhSachHaoPhi = haophis;
                }
            }
        }
        return dsCongTac;
    }

    /// <summary>
    /// Lưu định mức công tác mới (kèm hao phí) vào DB. Trả về true nếu thêm mới, false nếu cập nhật.
    /// </summary>
    public bool Insert(CongTacXayDung congTac)
    {
        using var connection = _context.GetConnection();
        using var transaction = connection.BeginTransaction();

        try
        {
            var sqlCheck = "SELECT Id FROM CongTacXayDung WHERE MaHieu = @MaHieu;";
            var existingId = connection.ExecuteScalar<int?>(sqlCheck, new { MaHieu = congTac.MaHieu }, transaction);

            int newId;
            bool isNew = false;
            if (existingId.HasValue)
            {
                newId = existingId.Value;
                var sqlUpdateCT = "UPDATE CongTacXayDung SET TenCongTac = @TenCongTac, DonVi = @DonVi WHERE Id = @Id;";
                connection.Execute(sqlUpdateCT, new { congTac.TenCongTac, congTac.DonVi, Id = newId }, transaction);

                var sqlDeleteHP = "DELETE FROM HaoPhi WHERE CongTacId = @Id;";
                connection.Execute(sqlDeleteHP, new { Id = newId }, transaction);
            }
            else
            {
                isNew = true;
                var sqlInsertCT = @"
                    INSERT INTO CongTacXayDung (MaHieu, TenCongTac, DonVi) 
                    VALUES (@MaHieu, @TenCongTac, @DonVi);
                    SELECT last_insert_rowid();";
                newId = connection.ExecuteScalar<int>(sqlInsertCT, congTac, transaction);
            }
            
            congTac.Id = newId;

            if (congTac.DanhSachHaoPhi.Any())
            {
                var sqlInsertHP = @"
                    INSERT INTO HaoPhi (CongTacId, LoaiHaoPhi, MaHieuHP, TenHaoPhi, DonVi, DinhMuc, HeSo) 
                    VALUES (@CongTacId, @LoaiHaoPhi, @MaHieuHP, @TenHaoPhi, @DonVi, @DinhMuc, @HeSo);";

                foreach (var hp in congTac.DanhSachHaoPhi)
                {
                    hp.CongTacId = newId;
                    connection.Execute(sqlInsertHP, hp, transaction);
                }
            }

            transaction.Commit();
            return isNew;
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }
}
