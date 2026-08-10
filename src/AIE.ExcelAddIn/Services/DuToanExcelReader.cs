using AIE.Core.Enums;
using AIE.Core.Models;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace AIE.ExcelAddIn.Services;

public class DuToanExcelReader
{
    public List<CongTacThamDinh> Read(ThamDinhConfig config)
    {
        var app = (Application)ExcelDnaUtil.Application;
        var wb = app.ActiveWorkbook;
        Worksheet ws = null;
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name == config.SheetName)
            {
                ws = sheet;
                break;
            }
        }

        if (ws == null) throw new Exception($"Không tìm thấy sheet '{config.SheetName}' trong file đang mở.");

        var result = new List<CongTacThamDinh>();
        int row = config.DongBatDau;

        Range usedRange = ws.UsedRange;
        int maxRow = usedRange.Rows.Count + usedRange.Row - 1;

        CongTacThamDinh currentCongTac = null;
        LoaiHaoPhi currentLoaiHp = LoaiHaoPhi.VL; // Mặc định là VL

        while (row <= maxRow)
        {
            string maHieu = GetCellValue(ws, row, config.ColMaHieu);
            string ten = GetCellValue(ws, row, config.ColTenCT);
            string donVi = GetCellValue(ws, row, config.ColDonVi);
            string strDinhMuc = GetCellValue(ws, row, config.ColDinhMuc);
            string strDonGia = string.IsNullOrEmpty(config.ColDonGia) ? "" : GetCellValue(ws, row, config.ColDonGia);
            string strThanhTien = string.IsNullOrEmpty(config.ColThanhTien) ? "" : GetCellValue(ws, row, config.ColThanhTien);

            // Bỏ qua dòng trống hoàn toàn ở các cột quan trọng
            if (string.IsNullOrEmpty(maHieu) && string.IsNullOrEmpty(ten) && string.IsNullOrEmpty(strDinhMuc) && string.IsNullOrEmpty(strDonGia) && string.IsNullOrEmpty(strThanhTien))
            {
                row++;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(maHieu))
            {
                currentCongTac = new CongTacThamDinh
                {
                    SoDongExcel = row,
                    MaHieu = maHieu,
                    TenCongTac = ten,
                    DonVi = donVi
                };
                result.Add(currentCongTac);
            }
            else if (currentCongTac != null)
            {
                // Kiểm tra xem dòng này có phải là dòng tiêu đề nhóm không (Vật liệu, Nhân công, Máy thi công)
                string tenLower = ten.ToLower().Trim();
                
                // Nếu dòng này không có định mức (hoặc định mức rỗng), nó có khả năng cao là dòng tiêu đề (VD: 'b) Nhân công')
                if (string.IsNullOrEmpty(strDinhMuc) || !decimal.TryParse(strDinhMuc, out _))
                {
                    if (tenLower.Contains("vật liệu") || tenLower == "vl")
                        currentLoaiHp = LoaiHaoPhi.VL;
                    else if (tenLower.Contains("nhân công") || tenLower == "nc")
                        currentLoaiHp = LoaiHaoPhi.NC;
                    else if (tenLower.Contains("máy thi công") || tenLower == "m" || tenLower.Contains("máy tc") || tenLower.Contains("máy thi cong"))
                        currentLoaiHp = LoaiHaoPhi.MAY;
                }
                else if (decimal.TryParse(strDinhMuc, out decimal dinhMuc))
                {
                    // Nếu là dòng dữ liệu (có định mức), ta có thể fallback thêm nếu tiêu đề bị thiếu
                    if (tenLower.StartsWith("nhân công")) currentLoaiHp = LoaiHaoPhi.NC;
                    else if (tenLower.StartsWith("máy")) currentLoaiHp = LoaiHaoPhi.MAY;

                    decimal donGia = 0;
                    if (!string.IsNullOrEmpty(strDonGia))
                    {
                        decimal.TryParse(strDonGia, out donGia);
                    }
                    
                    decimal thanhTien = 0;
                    if (!string.IsNullOrEmpty(strThanhTien))
                    {
                        decimal.TryParse(strThanhTien, out thanhTien);
                    }
                    
                    // Thêm HaoPhi
                    var hp = new HaoPhiThamDinh
                    {
                        SoDongExcel = row,
                        Loai = currentLoaiHp,
                        TenHaoPhi = ten,
                        DonVi = donVi,
                        DinhMuc = dinhMuc,
                        DonGia = donGia,
                        ThanhTien = thanhTien
                    };
                    currentCongTac.DanhSachHaoPhi.Add(hp);
                }
            }

            row++;
        }

        return result;
    }

    private string GetCellValue(Worksheet ws, int row, string col)
    {
        try
        {
            Range range = ws.Range[$"{col}{row}"];
            if (range.Value2 == null) return string.Empty;
            return range.Value2.ToString().Trim();
        }
        catch
        {
            return string.Empty;
        }
    }
}
