using System;
using System.IO;
using System.Linq;
using AIE.Data;
using AIE.Data.Repositories;
using Dapper;

namespace VerifyDb
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
            string dbPath = Path.Combine(appDataFolder, "aie_database.sqlite");
            
            var context = new AieDbContext(dbPath);
            context.InitializeDatabase();
            
            // Fix NhanCong Nhom values based on MaNC mapping
            using var conn = context.GetConnection();
            
            // N97789 = Nhóm 1, N97790 = Nhóm 2, N97791 = Nhóm 3, N97792 = Nhóm 4
            conn.Execute("UPDATE NhanCong SET Nhom = 1 WHERE MaNC = 'N97789'");
            conn.Execute("UPDATE NhanCong SET Nhom = 2 WHERE MaNC = 'N97790'");
            conn.Execute("UPDATE NhanCong SET Nhom = 3 WHERE MaNC = 'N97791'");
            conn.Execute("UPDATE NhanCong SET Nhom = 4 WHERE MaNC = 'N97792'");
            
            Console.WriteLine("Updated NhanCong Nhom values.");
            
            // Verify
            var ncRepo = new NhanCongRepository(context);
            var allNc = ncRepo.GetAll().ToList();
            Console.WriteLine("\n=== BANG NHAN CONG (after fix) ===");
            foreach (var nc in allNc)
            {
                Console.WriteLine($"  MaNC={nc.MaNC}, TenNC={nc.TenNC}, Nhom={nc.Nhom}, DonGia={nc.DonGia}");
            }
            
            Console.WriteLine("\n=== FIX NGUYEN GIA ===");
            var updatedRows = conn.Execute("UPDATE DinhMucCaMay_TT37 SET NguyenGia = NguyenGia * 1000 WHERE NguyenGia < 10000000");
            Console.WriteLine($"Updated {updatedRows} machines with NguyenGia * 1000.");

            // Verify lookup now works
            Console.WriteLine("\n=== TEST LOOKUP ===");
            var allDm = conn.Query<AIE.Core.Models.DinhMucCaMay_TT37>("SELECT * FROM DinhMucCaMay_TT37 LIMIT 3").ToList();
            foreach (var dm in allDm)
            {
                var matchNc = allNc.FirstOrDefault(x => x.Nhom == dm.NhomNhanCong);
                Console.WriteLine($"  MaMay={dm.MaMay}, NguyenGia={dm.NguyenGia}, NhomNC={dm.NhomNhanCong} => NC={matchNc?.TenNC ?? "NONE"}, DonGia={matchNc?.DonGia ?? 0}");
            }
        }
    }
}
