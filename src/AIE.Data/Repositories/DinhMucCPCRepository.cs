using System.Collections.Generic;
using System.Data;
using System.Linq;
using AIE.Core.Models;
using Dapper;

namespace AIE.Data.Repositories;

public class DinhMucCPCRepository
{
    private readonly IDbConnection _connection;

    public DinhMucCPCRepository(IDbConnection connection)
    {
        _connection = connection;
    }

    public List<DinhMucCPC> GetAll()
    {
        return _connection.Query<DinhMucCPC>("SELECT * FROM DinhMucCPC").ToList();
    }

    public List<DinhMucCPC> GetByLoaiCongTrinh(string loaiCongTrinh, string? phanLoaiPhu = null)
    {
        bool isNongNghiep = loaiCongTrinh == "Nông nghiệp & PTNT" || loaiCongTrinh == "Nông nghiệp và môi trường";
        string sql = isNongNghiep 
            ? "SELECT * FROM DinhMucCPC WHERE LoaiCongTrinh IN ('Nông nghiệp & PTNT', 'Nông nghiệp và môi trường')"
            : "SELECT * FROM DinhMucCPC WHERE LoaiCongTrinh = @LoaiCongTrinh";

        if (phanLoaiPhu != null)
        {
            sql += " AND PhanLoaiPhu = @PhanLoaiPhu";
        }
        else
        {
            sql += " AND PhanLoaiPhu IS NULL";
        }

        return _connection.Query<DinhMucCPC>(sql, new { LoaiCongTrinh = loaiCongTrinh, PhanLoaiPhu = phanLoaiPhu }).ToList();
    }
}
