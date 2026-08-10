using System.Data;
using System.Data.SQLite;
using System.IO;

namespace AIE.Data
{
    public class AieDbContext
    {
        private readonly string _connectionString;

        public AieDbContext(string dbPath)
        {
            _connectionString = $"Data Source={dbPath};Version=3;";
        }

        public IDbConnection GetConnection()
        {
            var conn = new SQLiteConnection(_connectionString);
            conn.Open();
            return conn;
        }

        public void InitializeDatabase()
        {
            using var connection = GetConnection();
            using var command = connection.CreateCommand();

            // Bảng định mức công tác
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS CongTacXayDung (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    MaHieu TEXT NOT NULL UNIQUE,
                    TenCongTac TEXT NOT NULL,
                    DonVi TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng định mức hao phí (VL, NC, Máy)
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS HaoPhi (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    CongTacId INTEGER NOT NULL,
                    LoaiHaoPhi INTEGER NOT NULL,
                    MaHieuHP TEXT NOT NULL,
                    TenHaoPhi TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DinhMuc REAL NOT NULL,
                    HeSo REAL DEFAULT 1.0,
                    FOREIGN KEY(CongTacId) REFERENCES CongTacXayDung(Id)
                );
            ";
            command.ExecuteNonQuery();

            // Bảng giá Vật Liệu
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS VatLieu (
                    MaVL TEXT PRIMARY KEY,
                    TenVL TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DonGia REAL NOT NULL,
                    NhaSanXuat TEXT,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng giá Nhân Công
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS NhanCong (
                    MaNC TEXT PRIMARY KEY,
                    TenNC TEXT NOT NULL,
                    Nhom INTEGER NOT NULL,
                    DonVi TEXT NOT NULL,
                    DonGia REAL NOT NULL,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng giá Máy Thi Công
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS MayThiCong (
                    MaMay TEXT PRIMARY KEY,
                    TenMay TEXT NOT NULL,
                    DonVi TEXT NOT NULL,
                    DonGia REAL NOT NULL,
                    GhiChu TEXT,
                    NgayCapNhat TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng cấu hình Tỉ lệ phần trăm
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS TiLePhanTram (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiCongTrinh TEXT NOT NULL,
                    CapCongTrinh TEXT,
                    LoaiTiLe TEXT NOT NULL,
                    GiaTri REAL NOT NULL,
                    CoSoTinh TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();

            // Bảng định mức Tỉ lệ Tư vấn
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS DinhMucTuVan (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    LoaiTuVan TEXT NOT NULL,
                    LoaiCongTrinh TEXT NOT NULL,
                    QuyMoMin REAL NOT NULL,
                    QuyMoMax REAL,
                    TiLe REAL NOT NULL,
                    GhiChu TEXT
                );
            ";
            command.ExecuteNonQuery();

            // Bảng Từ điển đồng nghĩa
            command.CommandText = @"
                CREATE TABLE IF NOT EXISTS TuDienDongNghia (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    TenGoc TEXT NOT NULL,
                    TenChuan TEXT NOT NULL,
                    LoaiVatTu TEXT NOT NULL, -- VL, NC, MAY
                    NgayTao TEXT NOT NULL
                );
            ";
            command.ExecuteNonQuery();
        }
    }
}
