using ClosedXML.Excel;
using System.IO;

namespace AIE.Data.ImportExport;

/// <summary>
/// Trình tạo file Excel mẫu (Template) để người dùng nhập dữ liệu định mức, giá cả.
/// </summary>
public class ExcelTemplateGenerator
{
    /// <summary>
    /// Tạo 6 file Excel mẫu tại thư mục chỉ định.
    /// 1. DinhMucCongTac.xlsx
    /// 2. GiaVatLieu.xlsx
    /// 3. GiaNhanCong.xlsx
    /// 4. GiaMayThiCong.xlsx
    /// 5. TiLePhanTram.xlsx
    /// 6. DinhMucTuVan.xlsx
    /// </summary>
    public void GenerateAllTemplates(string outputDirectory)
    {
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        GenerateDinhMucTemplate(Path.Combine(outputDirectory, "1_DinhMucCongTac.xlsx"));
        GenerateGiaVatLieuTemplate(Path.Combine(outputDirectory, "2_GiaVatLieu.xlsx"));
        GenerateGiaNhanCongTemplate(Path.Combine(outputDirectory, "3_GiaNhanCong.xlsx"));
        GenerateGiaMayThiCongTemplate(Path.Combine(outputDirectory, "4_GiaMayThiCong.xlsx"));
        GenerateTiLePhanTramTemplate(Path.Combine(outputDirectory, "5_TiLePhanTram.xlsx"));
        GenerateDinhMucTuVanTemplate(Path.Combine(outputDirectory, "6_DinhMucTuVan.xlsx"));
    }

    private void GenerateDinhMucTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        
        // Sheet 1: Công tác
        var wsCongTac = wb.Worksheets.Add("CongTac");
        wsCongTac.Cell(1, 1).Value = "MaHieu";
        wsCongTac.Cell(1, 2).Value = "TenCongTac";
        wsCongTac.Cell(1, 3).Value = "DonVi";
        
        wsCongTac.Cell(2, 1).Value = "AF.11112";
        wsCongTac.Cell(2, 2).Value = "Bê tông lót móng, đá 4x6, M100";
        wsCongTac.Cell(2, 3).Value = "m3";

        // Sheet 2: Hao phí
        var wsHaoPhi = wb.Worksheets.Add("HaoPhi");
        wsHaoPhi.Cell(1, 1).Value = "MaHieuCongTac";
        wsHaoPhi.Cell(1, 2).Value = "LoaiHaoPhi"; // VL, NC, MAY
        wsHaoPhi.Cell(1, 3).Value = "MaHieuHP";
        wsHaoPhi.Cell(1, 4).Value = "TenHaoPhi";
        wsHaoPhi.Cell(1, 5).Value = "DonVi";
        wsHaoPhi.Cell(1, 6).Value = "DinhMuc";

        wsHaoPhi.Cell(2, 1).Value = "AF.11112";
        wsHaoPhi.Cell(2, 2).Value = "VL";
        wsHaoPhi.Cell(2, 3).Value = "V08770";
        wsHaoPhi.Cell(2, 4).Value = "Xi măng PCB40";
        wsHaoPhi.Cell(2, 5).Value = "kg";
        wsHaoPhi.Cell(2, 6).Value = 234.725;

        FormatHeader(wsCongTac);
        FormatHeader(wsHaoPhi);
        wb.SaveAs(filePath);
    }

    private void GenerateGiaVatLieuTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        var ws = wb.Worksheets.Add("GiaVatLieu");
        ws.Cell(1, 1).Value = "MaVL";
        ws.Cell(1, 2).Value = "TenVL";
        ws.Cell(1, 3).Value = "DonVi";
        ws.Cell(1, 4).Value = "DonGia";
        ws.Cell(1, 5).Value = "NhaSanXuat";
        ws.Cell(1, 6).Value = "GhiChu";

        ws.Cell(2, 1).Value = "V08770";
        ws.Cell(2, 2).Value = "Xi măng PCB40";
        ws.Cell(2, 3).Value = "kg";
        ws.Cell(2, 4).Value = 1741;

        FormatHeader(ws);
        wb.SaveAs(filePath);
    }

    private void GenerateGiaNhanCongTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        var ws = wb.Worksheets.Add("GiaNhanCong");
        ws.Cell(1, 1).Value = "MaNC";
        ws.Cell(1, 2).Value = "TenNC";
        ws.Cell(1, 3).Value = "Nhom";
        ws.Cell(1, 4).Value = "DonVi";
        ws.Cell(1, 5).Value = "DonGia";
        ws.Cell(1, 6).Value = "GhiChu";

        ws.Cell(2, 1).Value = "N97790";
        ws.Cell(2, 2).Value = "Nhân công nhóm 2";
        ws.Cell(2, 3).Value = 2;
        ws.Cell(2, 4).Value = "công";
        ws.Cell(2, 5).Value = 0; // Điền sau

        FormatHeader(ws);
        wb.SaveAs(filePath);
    }

    private void GenerateGiaMayThiCongTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        var ws = wb.Worksheets.Add("GiaMayThiCong");
        ws.Cell(1, 1).Value = "MaMay";
        ws.Cell(1, 2).Value = "TenMay";
        ws.Cell(1, 3).Value = "DonVi";
        ws.Cell(1, 4).Value = "DonGia";
        ws.Cell(1, 5).Value = "GhiChu";

        ws.Cell(2, 1).Value = "M112.1101";
        ws.Cell(2, 2).Value = "Máy đầm bê tông";
        ws.Cell(2, 3).Value = "ca";
        ws.Cell(2, 4).Value = 0;

        FormatHeader(ws);
        wb.SaveAs(filePath);
    }

    private void GenerateTiLePhanTramTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        var ws = wb.Worksheets.Add("TiLePhanTram");
        ws.Cell(1, 1).Value = "LoaiCongTrinh";
        ws.Cell(1, 2).Value = "CapCongTrinh";
        ws.Cell(1, 3).Value = "LoaiTiLe"; // CPC, TT, TNCTTT, GTGT, NhaTam
        ws.Cell(1, 4).Value = "GiaTri";
        ws.Cell(1, 5).Value = "CoSoTinh"; // T, NC, T+GT, G, Gxd

        ws.Cell(2, 1).Value = "DanDung";
        ws.Cell(2, 2).Value = "";
        ws.Cell(2, 3).Value = "CPC";
        ws.Cell(2, 4).Value = 7.3;
        ws.Cell(2, 5).Value = "T";

        FormatHeader(ws);
        wb.SaveAs(filePath);
    }

    private void GenerateDinhMucTuVanTemplate(string filePath)
    {
        using var wb = CreateTemplateWorkbook();
        var ws = wb.Worksheets.Add("DinhMucTuVan");
        ws.Cell(1, 1).Value = "LoaiTuVan";
        ws.Cell(1, 2).Value = "LoaiCongTrinh";
        ws.Cell(1, 3).Value = "QuyMoMin";
        ws.Cell(1, 4).Value = "QuyMoMax";
        ws.Cell(1, 5).Value = "TiLe";
        ws.Cell(1, 6).Value = "GhiChu";

        ws.Cell(2, 1).Value = "ThietKe";
        ws.Cell(2, 2).Value = "DanDung";
        ws.Cell(2, 3).Value = 15;
        ws.Cell(2, 4).Value = 50;
        ws.Cell(2, 5).Value = 3.5;

        FormatHeader(ws);
        wb.SaveAs(filePath);
    }

    private void FormatHeader(IXLWorksheet ws)
    {
        var header = ws.Row(1);
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.LightGray;
        ws.Columns().AdjustToContents();
    }

    private XLWorkbook CreateTemplateWorkbook()
    {
        var wb = new XLWorkbook();
        wb.Style.Font.FontName = "Times New Roman";
        wb.Style.Font.FontSize = 11;
        return wb;
    }
}
