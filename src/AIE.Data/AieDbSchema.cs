namespace AIE.Data;

/// <summary>
/// Scripts tạo cấu trúc bảng cho CSDL SQLite nội bộ (AIE_Data.sqlite).
/// </summary>
public static class AieDbSchema
{
    public const string CreateTablesSql = @"
        -- 1. Bảng Công tác xây dựng
        CREATE TABLE IF NOT EXISTS CongTacXayDung (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            MaHieu TEXT NOT NULL UNIQUE,
            TenCongTac TEXT NOT NULL,
            DonVi TEXT NOT NULL
        );

        -- 2. Bảng Hao phí (Mapping 1-N với CongTacXayDung)
        CREATE TABLE IF NOT EXISTS HaoPhi (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            CongTacId INTEGER NOT NULL,
            LoaiHaoPhi INTEGER NOT NULL, -- 0: VL, 1: NC, 2: MAY (Enum LoaiHaoPhi)
            MaHieuHP TEXT NOT NULL,
            TenHaoPhi TEXT NOT NULL,
            DonVi TEXT NOT NULL,
            DinhMuc REAL NOT NULL,
            HeSo REAL DEFAULT 1.0,
            FOREIGN KEY(CongTacId) REFERENCES CongTacXayDung(Id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS IDX_HaoPhi_CongTacId ON HaoPhi(CongTacId);

        -- 3. Bảng Vật liệu (Giá địa phương Đà Nẵng)
        CREATE TABLE IF NOT EXISTS VatLieu (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            MaVL TEXT NOT NULL UNIQUE,
            TenVL TEXT NOT NULL,
            DonVi TEXT NOT NULL,
            DonGia REAL NOT NULL DEFAULT 0,
            NhaSanXuat TEXT,
            GhiChu TEXT,
            NgayCapNhat TEXT NOT NULL
        );

        -- 4. Bảng Nhân công (Giá địa phương Đà Nẵng)
        CREATE TABLE IF NOT EXISTS NhanCong (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            MaNC TEXT NOT NULL UNIQUE,
            TenNC TEXT NOT NULL,
            Nhom INTEGER NOT NULL, -- Nhóm 1-6
            DonVi TEXT NOT NULL DEFAULT 'công',
            DonGia REAL NOT NULL DEFAULT 0,
            GhiChu TEXT,
            NgayCapNhat TEXT NOT NULL
        );

        -- 5. Bảng Máy thi công (Giá địa phương Đà Nẵng)
        CREATE TABLE IF NOT EXISTS MayThiCong (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            MaMay TEXT NOT NULL UNIQUE,
            TenMay TEXT NOT NULL,
            DonVi TEXT NOT NULL DEFAULT 'ca',
            DonGia REAL NOT NULL DEFAULT 0,
            GhiChu TEXT,
            NgayCapNhat TEXT NOT NULL
        );

        -- 6. Bảng Tỉ lệ phần trăm (CPC, TT, TNCTTT,...) theo loại công trình
        CREATE TABLE IF NOT EXISTS TiLePhanTram (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            LoaiCongTrinh TEXT NOT NULL,
            CapCongTrinh TEXT,
            LoaiTiLe TEXT NOT NULL,
            GiaTri REAL NOT NULL,
            CoSoTinh TEXT NOT NULL
        );

        -- 7. Bảng Định mức % Quản lý dự án
        CREATE TABLE IF NOT EXISTS DinhMucQLDA (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            LoaiDuAn TEXT NOT NULL,
            QuyMoMin REAL NOT NULL,
            QuyMoMax REAL,
            TiLe REAL NOT NULL,
            GhiChu TEXT
        );

        -- 8. Bảng Định mức % Tư vấn
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
}
