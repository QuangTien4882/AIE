using AIE.Data.ImportExport;
using AIE.Data.Repositories;
using System;
using System.IO;

namespace AIE.Data
{
    /// <summary>
    /// Quản lý tập trung database SQLite: khởi tạo, import dữ liệu.
    /// </summary>
    public class DatabaseManager
    {
        private readonly string _dbPath;
        private AieDbContext _context;

        public DatabaseManager()
        {
            // Lưu DB tại thư mục AppData của người dùng
            var appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AIE_DuToan");
            if (!Directory.Exists(appDataFolder))
                Directory.CreateDirectory(appDataFolder);

            _dbPath = Path.Combine(appDataFolder, "aie_database.sqlite");
            _context = new AieDbContext(_dbPath);
            _context.InitializeDatabase();
        }

        public string DbPath { get { return _dbPath; } }
        public AieDbContext Context { get { return _context; } }

        /// <summary>
        /// Import file 1_DinhMucCongTac.xlsx
        /// </summary>
        public ImportResult ImportDinhMucCongTac(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var danhSachCT = importService.ImportDinhMucCongTac(filePath, out result);

            var repo = new CongTacRepository(_context);
            int newCount = 0;
            int updateCount = 0;
            foreach (var ct in danhSachCT)
            {
                if (repo.Insert(ct))
                    newCount++;
                else
                    updateCount++;
            }
            
            result.SoLuongMoi += newCount;
            result.SoLuongCapNhat += updateCount;

            return result;
        }

        /// <summary>
        /// Import trực tiếp từ file F1 xuất ra
        /// </summary>
        public ImportResult ImportF1DinhMucCongTac(string filePath)
        {
            var importer = new F1DatabaseImporter();
            var danhSachCT = importer.ParseF1File(filePath);

            var repo = new CongTacRepository(_context);
            int countCT = 0;
            int countHP = 0;
            int newCount = 0;
            int updateCount = 0;
            
            foreach (var ct in danhSachCT)
            {
                if (repo.Insert(ct))
                    newCount++;
                else
                    updateCount++;
                    
                countCT++;
                countHP += ct.DanhSachHaoPhi.Count;
            }

            var result = new ImportResult();
            result.SoLuongThanhCong = countCT + countHP;
            result.SoLuongMoi = newCount;
            result.SoLuongCapNhat = updateCount;
            return result;
        }

        /// <summary>
        /// Import file 2_GiaVatLieu.xlsx
        /// </summary>
        public ImportResult ImportGiaVatLieu(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaVatLieu(filePath, out result);

            var repo = new VatLieuRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        /// <summary>
        /// Import file 3_GiaNhanCong.xlsx
        /// </summary>
        public ImportResult ImportGiaNhanCong(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaNhanCong(filePath, out result);

            var repo = new NhanCongRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        /// <summary>
        /// Import file 4_GiaMayThiCong.xlsx
        /// </summary>
        public ImportResult ImportGiaMayThiCong(string filePath)
        {
            var importService = new ExcelImportService();
            ImportResult result;
            var data = importService.ImportGiaMayThiCong(filePath, out result);

            var repo = new MayThiCongRepository(_context);
            foreach (var item in data)
            {
                repo.Upsert(item);
            }

            return result;
        }

        public void ClearAllDinhMuc()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM HaoPhi;
                    DELETE FROM CongTacXayDung;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'HaoPhi' OR name = 'CongTacXayDung';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllVatLieu()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM VatLieu;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'VatLieu';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllNhanCong()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM NhanCong;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'NhanCong';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        public void ClearAllMayThiCong()
        {
            using (var connection = _context.GetConnection())
            {
                var sql = @"
                    DELETE FROM MayThiCong;
                    UPDATE sqlite_sequence SET seq = 0 WHERE name = 'MayThiCong';
                ";
                Dapper.SqlMapper.Execute(connection, sql);
            }
        }

        /// <summary>
        /// Lấy thống kê số lượng dữ liệu trong DB.
        /// </summary>
        public DatabaseStats GetStats()
        {
            var stats = new DatabaseStats();
            var vlRepo = new VatLieuRepository(_context);
            var ncRepo = new NhanCongRepository(_context);
            var mayRepo = new MayThiCongRepository(_context);
            var ctRepo = new CongTacRepository(_context);

            stats.SoCongTac = ctRepo.Count();
            stats.SoVatLieu = vlRepo.GetAll().ToList().Count;
            stats.SoNhanCong = ncRepo.GetAll().ToList().Count;
            stats.SoMayThiCong = mayRepo.GetAll().ToList().Count;

            return stats;
        }
    }

    public class DatabaseStats
    {
        public int SoCongTac { get; set; }
        public int SoVatLieu { get; set; }
        public int SoNhanCong { get; set; }
        public int SoMayThiCong { get; set; }

        public string TomTat
        {
            get
            {
                return string.Format(
                    "Công tác: {0}\nVật liệu: {1}\nNhân công: {2}\nMáy thi công: {3}",
                    SoCongTac, SoVatLieu, SoNhanCong, SoMayThiCong);
            }
        }
    }
}
