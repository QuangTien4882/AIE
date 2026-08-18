using System.Collections.Generic;
using System.Data;
using System.Linq;
using AIE.Core.Models;
using Dapper;

namespace AIE.Data.Repositories;

public class DinhMucTTRepository
{
    private readonly IDbConnection _connection;

    public DinhMucTTRepository(IDbConnection connection)
    {
        _connection = connection;
    }

    public List<DinhMucTT> GetAll()
    {
        return _connection.Query<DinhMucTT>("SELECT * FROM DinhMucTT").ToList();
    }

    public DinhMucTT? GetByLoaiCongTrinh(string loaiCongTrinh, string? phanLoaiPhu = null)
    {
        string sql = "SELECT * FROM DinhMucTT WHERE LoaiCongTrinh = @LoaiCongTrinh";
        if (phanLoaiPhu != null)
        {
            sql += " AND PhanLoaiPhu = @PhanLoaiPhu";
        }
        else
        {
            sql += " AND PhanLoaiPhu IS NULL";
        }

        return _connection.QueryFirstOrDefault<DinhMucTT>(sql, new { LoaiCongTrinh = loaiCongTrinh, PhanLoaiPhu = phanLoaiPhu });
    }
}
