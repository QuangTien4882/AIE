using System;
using System.IO;
using System.Linq;
using Dapper;
using System.Data.SQLite;

namespace VerifyDb
{
    class Program
    {
        static void Main()
        {
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
            string dbPath = Path.Combine(appDataFolder, "aie_database.sqlite");
            using var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            conn.Open();

            var list = conn.Query("SELECT Nhom, LoaiNhanCong, TenNC, DonGiaVung2 FROM NhanCong").ToList();
            Console.WriteLine($"Total rows: {list.Count}");
            foreach(var item in list) {
                Console.WriteLine($"Nhom: {item.Nhom}, Loai: {item.LoaiNhanCong}, Ten: {item.TenNC}, Gia: {item.DonGiaVung2}");
            }
        }
    }
}
