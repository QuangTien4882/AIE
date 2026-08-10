using Xunit;
using AIE.Core.Models;
using AIE.Core.Services.LapDuToan;
using AIE.Core.Enums;
using System.Collections.Generic;

namespace AIE.Core.Tests;

public class PhanTichDonGiaTests
{
    [Fact]
    public void PhanTichDonGia_TheoDuLieu_Screenshot()
    {
        // Arrange
        var service = new PhanTichDonGiaService();
        var congTac = new CongTacXayDung
        {
            MaHieu = "AF.11112",
            TenCongTac = "Bê tông lót móng",
            DonVi = "m3",
            DanhSachHaoPhi = new List<HaoPhi>
            {
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.VL, MaHieuHP = "V08770", TenHaoPhi = "Xi măng PCB40", DonVi = "kg", DinhMuc = 234.725m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.VL, MaHieuHP = "V00112", TenHaoPhi = "Cát vàng", DonVi = "m3", DinhMuc = 0.56375m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.VL, MaHieuHP = "V05209", TenHaoPhi = "Đá 4x6", DonVi = "m3", DinhMuc = 0.915325m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.VL, MaHieuHP = "V00494", TenHaoPhi = "Nước", DonVi = "lít", DinhMuc = 166.05m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.NC, MaHieuHP = "N97790", TenHaoPhi = "Nhân công nhóm 2", DonVi = "công", DinhMuc = 1.07m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.MAY, MaHieuHP = "M112.1101", TenHaoPhi = "Máy đầm bê tông", DonVi = "ca", DinhMuc = 0.089m },
                new HaoPhi { LoaiHaoPhi = LoaiHaoPhi.MAY, MaHieuHP = "M104.0102", TenHaoPhi = "Máy trộn bê tông", DonVi = "ca", DinhMuc = 0.095m }
            }
        };

        var giaVL = new Dictionary<string, decimal>
        {
            { "V08770", 1741m },
            { "V00112", 250000m },
            { "V05209", 230909m },
            { "V00494", 7m }
        };
        var giaNC = new Dictionary<string, decimal>(); // Không có giá trong screenshot
        var giaMay = new Dictionary<string, decimal>(); // Không có giá trong screenshot

        // Act
        var ketQua = service.PhanTich(congTac, giaVL, giaNC, giaMay);

        // Assert
        Assert.Equal(408656.225m, ketQua.ChiTietVatLieu[0].ThanhTien);
        Assert.Equal(140937.5m, ketQua.ChiTietVatLieu[1].ThanhTien);
        Assert.Equal(211356.780425m, ketQua.ChiTietVatLieu[2].ThanhTien);
        Assert.Equal(1162.35m, ketQua.ChiTietVatLieu[3].ThanhTien);

        // Làm tròn phần nguyên giống phần mềm dự toán
        decimal tongVL = Math.Round(ketQua.DonGiaVL, 0, MidpointRounding.AwayFromZero);
        Assert.Equal(762113m, tongVL); // Khớp với 762.113 trong screenshot
        
        Assert.Equal(0m, ketQua.DonGiaNC);
        Assert.Equal(0m, ketQua.DonGiaMay);
    }
}
