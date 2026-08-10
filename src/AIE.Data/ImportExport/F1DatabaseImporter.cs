using AIE.Core.Models;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace AIE.Data.ImportExport;

/// <summary>
/// Đọc trực tiếp file Hao phí định mức xuất từ phần mềm F1
/// </summary>
public class F1DatabaseImporter
{
    public List<CongTacXayDung> ParseF1File(string filePath)
    {
        var listCongTac = new List<CongTacXayDung>();
        
        using var wb = new XLWorkbook(filePath);
        var ws = wb.Worksheet(1); // Lấy sheet đầu tiên
        
        var rows = ws.RowsUsed();
        
        CongTacXayDung? currentCongTac = null;

        foreach (var row in rows)
        {
            // Bỏ qua các dòng tiêu đề (Dòng có STT, Mã hiệu, Tên công tác...)
            if (row.Cell(2).GetString().Equals("Mã hiệu", StringComparison.OrdinalIgnoreCase))
                continue;

            string maHieu = row.Cell(2).GetString().Trim();
            string ten = row.Cell(3).GetString().Trim();
            string donVi = row.Cell(4).GetString().Trim();
            
            // Lấy mức hao phí
            string hpVatLieu = row.Cell(6).GetString().Trim();
            string hpNhanCong = row.Cell(7).GetString().Trim();
            string hpMayThiCong = row.Cell(8).GetString().Trim();

            // Nếu cột Mã hiệu có giá trị, và các cột Hao phí ĐỀU RỖNG -> Đây là dòng CÔNG TÁC
            if (!string.IsNullOrEmpty(maHieu) && string.IsNullOrEmpty(hpVatLieu) && string.IsNullOrEmpty(hpNhanCong) && string.IsNullOrEmpty(hpMayThiCong))
            {
                currentCongTac = new CongTacXayDung
                {
                    MaHieu = maHieu,
                    TenCongTac = ten,
                    DonVi = donVi,
                    DanhSachHaoPhi = new List<HaoPhi>()
                };
                listCongTac.Add(currentCongTac);
            }
            // Nếu cột Mã hiệu có giá trị, và CÓ ít nhất 1 cột Hao phí -> Đây là dòng HAO PHÍ (Vật liệu, nhân công hoặc máy)
            else if (currentCongTac != null && !string.IsNullOrEmpty(maHieu))
            {
                var hp = new HaoPhi
                {
                    MaHieuHP = maHieu,
                    TenHaoPhi = ten,
                    DonVi = donVi,
                    HeSo = 1.0M
                };

                // Xác định loại hao phí dựa trên cột nào có số liệu
                if (decimal.TryParse(hpVatLieu, out decimal klVl))
                {
                    hp.LoaiHaoPhi = AIE.Core.Enums.LoaiHaoPhi.VL;
                    hp.DinhMuc = klVl;
                    currentCongTac.DanhSachHaoPhi.Add(hp);
                }
                else if (decimal.TryParse(hpNhanCong, out decimal klNc))
                {
                    hp.LoaiHaoPhi = AIE.Core.Enums.LoaiHaoPhi.NC;
                    hp.DinhMuc = klNc;
                    currentCongTac.DanhSachHaoPhi.Add(hp);
                }
                else if (decimal.TryParse(hpMayThiCong, out decimal klMtc))
                {
                    hp.LoaiHaoPhi = AIE.Core.Enums.LoaiHaoPhi.MAY;
                    hp.DinhMuc = klMtc;
                    currentCongTac.DanhSachHaoPhi.Add(hp);
                }
            }
        }

        return listCongTac;
    }
}
