using System;
using System.Data.SQLite;
using System.IO;

class Program
{
    static void Main()
    {
        string appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
        string dbPath = Path.Combine(appData, "aie_database.sqlite");
        
        if (!File.Exists(dbPath))
        {
            Console.WriteLine("DB not found at " + dbPath);
            return;
        }

        string connectionString = $"Data Source={dbPath};Version=3;";
        string outputMd = @"C:\Users\quang\.gemini\antigravity-ide\brain\96f85643-e6e2-4978-935e-6c0842c38a46\db_verification.md";
        
        using (StreamWriter sw = new StreamWriter(outputMd))
        {
            sw.WriteLine("# Dữ liệu Định mức (Trích xuất 5 công tác đầu tiên)");
            sw.WriteLine();
            
            using (var conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                
                // Get count
                using (var cmdCount = new SQLiteCommand("SELECT COUNT(*) FROM CongTacXayDung", conn))
                {
                    var count = cmdCount.ExecuteScalar();
                    sw.WriteLine($"> **Tổng số Công tác đang có trong DB:** {count}");
                }

                using (var cmdCountHP = new SQLiteCommand("SELECT COUNT(*) FROM HaoPhi", conn))
                {
                    var count = cmdCountHP.ExecuteScalar();
                    sw.WriteLine($"> **Tổng số chi tiết Hao phí đang có trong DB:** {count}");
                }
                sw.WriteLine();

                // Get top 5
                using (var cmd = new SQLiteCommand("SELECT Id, MaHieu, TenCongTac, DonVi FROM CongTacXayDung LIMIT 5", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int id = reader.GetInt32(0);
                        string maHieu = reader.GetString(1);
                        string ten = reader.GetString(2);
                        string dv = reader.GetString(3);
                        
                        sw.WriteLine($"### {maHieu} - {ten} ({dv})");
                        sw.WriteLine("| Loại | Mã Vật Tư | Tên Vật Tư | Đơn vị | Định mức |");
                        sw.WriteLine("|---|---|---|---|---|");

                        using (var cmdHP = new SQLiteCommand($"SELECT LoaiHaoPhi, MaHieuHP, TenHaoPhi, DonVi, DinhMuc FROM HaoPhi WHERE CongTacId = {id}", conn))
                        using (var readerHP = cmdHP.ExecuteReader())
                        {
                            while (readerHP.Read())
                            {
                                int loai = readerHP.GetInt32(0);
                                string maHP = readerHP.GetString(1);
                                string tenHP = readerHP.GetString(2);
                                string dvHP = readerHP.GetString(3);
                                decimal dm = readerHP.GetDecimal(4);
                                
                                string loaiStr = loai == 0 ? "Vật liệu" : (loai == 1 ? "Nhân công" : "Máy TC");
                                sw.WriteLine($"| {loaiStr} | {maHP} | {tenHP} | {dvHP} | {dm} |");
                            }
                        }
                        sw.WriteLine();
                    }
                }
            }
        }
    }
}
