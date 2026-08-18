using AIE.Core.Models;
using Newtonsoft.Json;
using System;
using System.IO;

namespace AIE.ExcelAddIn.Services
{
    /// <summary>
    /// Dịch vụ lưu/đọc file dự toán (.dt).
    /// File .dt bản chất là JSON chứa toàn bộ đối tượng DuToan.
    /// </summary>
    public static class DuToanFileService
    {
        private static readonly JsonSerializerSettings _settings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore,
            TypeNameHandling = TypeNameHandling.Auto, // Cần thiết để deserialize đúng kiểu kế thừa (VatLieuHienTruong, NhanCongHienTruong, MayThiCongHienTruong)
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
        };

        /// <summary>
        /// Lưu đối tượng DuToan ra file .dt (JSON)
        /// </summary>
        public static void Save(DuToan duToan, string filePath)
        {
            if (duToan == null) throw new ArgumentNullException(nameof(duToan));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("Đường dẫn file không hợp lệ.", nameof(filePath));

            var json = JsonConvert.SerializeObject(duToan, _settings);
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// Đọc file .dt và trả về đối tượng DuToan
        /// </summary>
        public static DuToan Load(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Không tìm thấy file dự toán.", filePath);

            var json = File.ReadAllText(filePath);
            var duToan = JsonConvert.DeserializeObject<DuToan>(json, _settings);

            if (duToan == null)
                throw new InvalidOperationException("Không thể đọc dữ liệu dự toán từ file.");

            return duToan;
        }
    }
}
