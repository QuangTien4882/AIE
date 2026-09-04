using System;
using System.Collections.Generic;
using System.Linq;
using AIE.Core.Models;
using Xunit;

namespace AIE.Core.Tests;

public class VanChuyenTests
{
    [Fact]
    public void TinhHaoPhiCaXeOTo_ChuanTheoViDuThongTu12_Trang686()
    {
        // Ví dụ trong Thông tư số 12/2021/TT-BXD (trang 686):
        // Vận chuyển cát bằng ô tô tự đổ cự ly 19km:
        // - 0,3km đầu đường loại 5 (k=1.50)
        // - 5km tiếp theo đường loại 3 (k=1.00)
        // - 2km tiếp theo đường loại 4 (k=1.35)
        // - 7km tiếp theo đường loại 2 (k=0.68)
        // - 3km tiếp theo đường loại 1 (k=0.57)
        // - 1,7km tiếp theo đường loại 3 (k=1.00)
        //
        // Mức hao phí ô tô tự đổ vận chuyển 19km =
        // Đm1 x (0.3*k5 + 0.7*k3)
        // + Đm2 x (4.3*k3 + 2*k4 + 2.7*k2)
        // + Đm3 x (4.3*k2 + 3*k1 + 1.7*k3)

        var cungDuongs = new List<CungDuongVanChuyen>
        {
            new() { TenDoanDuong = "Đoạn 1", CuLyKm = 0.3m, LoaiDuong = 5 },
            new() { TenDoanDuong = "Đoạn 2", CuLyKm = 5.0m, LoaiDuong = 3 },
            new() { TenDoanDuong = "Đoạn 3", CuLyKm = 2.0m, LoaiDuong = 4 },
            new() { TenDoanDuong = "Đoạn 4", CuLyKm = 7.0m, LoaiDuong = 2 },
            new() { TenDoanDuong = "Đoạn 5", CuLyKm = 3.0m, LoaiDuong = 1 },
            new() { TenDoanDuong = "Đoạn 6", CuLyKm = 1.7m, LoaiDuong = 3 },
        };

        var dmCat7t = DinhMucVanChuyenDatabase.DanhSachOTo.First(x => x.MaHieu == "AM.2311");
        // AM.2311: Dm1 = 0.027, Dm2 = 0.019, Dm3 = 0.014

        var (caXe, q1, q2, q3, q4) = DinhMucVanChuyenDatabase.TinhHaoPhiCaXeOTo(dmCat7t, cungDuongs);

        // Kiểm tra Nấc 1: 0.3*1.5 + 0.7*1.0 = 0.45 + 0.7 = 1.15
        decimal expectedQ1 = 0.3m * 1.50m + 0.7m * 1.00m;
        Assert.Equal(expectedQ1, q1);

        // Kiểm tra Nấc 2: 4.3*1.0 + 2.0*1.35 + 2.7*0.68 = 4.3 + 2.7 + 1.836 = 8.836
        decimal expectedQ2 = 4.3m * 1.00m + 2.0m * 1.35m + 2.7m * 0.68m;
        Assert.Equal(expectedQ2, q2);

        // Kiểm tra Nấc 3: 4.3*0.68 + 3.0*0.57 + 1.7*1.00 = 2.924 + 1.71 + 1.7 = 6.334
        decimal expectedQ3 = 4.3m * 0.68m + 3.0m * 0.57m + 1.7m * 1.00m;
        Assert.Equal(expectedQ3, q3);

        // Kiểm tra Nấc 4: cự ly 19km <= 60km nên q4 = 0
        Assert.Equal(0m, q4);

        // Tổng ca xe
        decimal expectedCaXe = dmCat7t.Dm1 * expectedQ1 + dmCat7t.Dm2 * expectedQ2 + dmCat7t.Dm3 * expectedQ3;
        Assert.Equal(expectedCaXe, caXe);
    }

    [Fact]
    public void TinhHaoPhiCaXeOTo_CuLyLonHon60Km_ApDungHeSo095ChoNac4()
    {
        var cungDuongs = new List<CungDuongVanChuyen>
        {
            new() { TenDoanDuong = "Tuyến đường dài", CuLyKm = 100m, LoaiDuong = 3 } // k=1.00
        };

        var dmCat7t = DinhMucVanChuyenDatabase.DanhSachOTo.First(x => x.MaHieu == "AM.2311");
        var (caXe, q1, q2, q3, q4) = DinhMucVanChuyenDatabase.TinhHaoPhiCaXeOTo(dmCat7t, cungDuongs);

        Assert.Equal(1.0m, q1);
        Assert.Equal(9.0m, q2);
        Assert.Equal(50.0m, q3);
        Assert.Equal(40.0m, q4);

        decimal expectedCa = dmCat7t.Dm1 * 1.0m + dmCat7t.Dm2 * 9.0m + dmCat7t.Dm3 * 50.0m + (dmCat7t.Dm3 * 0.95m) * 40.0m;
        Assert.Equal(expectedCa, caXe);
    }

    [Fact]
    public void TinhHaoPhiNhanCongBo_ChuanTheoDinhMucAM21000()
    {
        var dmCat = DinhMucVanChuyenDatabase.DanhSachBo.First(x => x.MaHieu == "AM.2101");
        // AM.2101: Dm10m = 0.075, DmTiepTheo = 0.008

        // Cự ly <= 10m: tính tròn 10m
        decimal haoPhi10m = DinhMucVanChuyenDatabase.TinhHaoPhiNhanCongBo(dmCat, 8m);
        Assert.Equal(0.075m, haoPhi10m);

        // Cự ly 30m: 10m đầu + 2 đoạn 10m tiếp theo
        decimal haoPhi30m = DinhMucVanChuyenDatabase.TinhHaoPhiNhanCongBo(dmCat, 30m);
        Assert.Equal(0.075m + 2 * 0.008m, haoPhi30m);

        // Cự ly 30m đường dốc k=1.35
        decimal haoPhiDoc = DinhMucVanChuyenDatabase.TinhHaoPhiNhanCongBo(dmCat, 30m, 1.35m);
        Assert.Equal((0.075m + 2 * 0.008m) * 1.35m, haoPhiDoc);

        // Cự ly 30m tầng 3 (tăng 2 tầng, mỗi tầng * 1.1)
        decimal haoPhiTang = DinhMucVanChuyenDatabase.TinhHaoPhiNhanCongBo(dmCat, 30m, 1.0m, 3);
        decimal expectedTang = (0.075m + 2 * 0.008m) * (decimal)Math.Pow(1.1, 2);
        Assert.Equal(expectedTang, haoPhiTang);
    }

    [Fact]
    public void NhanDienVatLieu_ChinhXac()
    {
        var oToCat = DinhMucVanChuyenDatabase.NhanDienOTo("Cát vàng hạt trung");
        Assert.NotNull(oToCat);
        Assert.Equal("AM.2311", oToCat.MaHieu);

        var oToThep = DinhMucVanChuyenDatabase.NhanDienOTo("Thép thanh vằn D16");
        Assert.NotNull(oToThep);
        Assert.Equal("AM.2451", oToThep.MaHieu);

        var oToCong = DinhMucVanChuyenDatabase.NhanDienOTo("Cống bê tông ly tâm D800");
        Assert.NotNull(oToCong);
        Assert.Equal("AM.2611", oToCong.MaHieu);

        var boGach = DinhMucVanChuyenDatabase.NhanDienBo("Gạch đất sét nung 6 lỗ");
        Assert.NotNull(boGach);
        Assert.Equal("AM.2105", boGach.MaHieu);
    }
}
