using Xunit;
using AIE.Core.Services.LapDuToan;

namespace AIE.Core.Tests;

public class ChiPhiXayDungTests
{
    [Fact]
    public void ChiPhiXayDungCalc_TheoDuLieu_Screenshot()
    {
        // Arrange
        var calc = new ChiPhiXayDungCalc();
        decimal tongVL = 762113m;
        decimal tongNC = 0m;
        decimal tongMay = 0m;
        
        decimal tiLeCPC = 7.3m;
        decimal tiLeTT = 2.5m;
        decimal tiLeTNCTTT = 5.5m;
        decimal tiLeGTGT = 8.0m;
        decimal tiLeNhaTam = 1.1m;

        // Act
        var result = calc.Tinh(tongVL, tongNC, tongMay, tiLeCPC, tiLeTT, tiLeTNCTTT, tiLeGTGT, tiLeNhaTam);

        // Assert
        Assert.Equal(762113m, result.T);
        Assert.Equal(55634m, result.CPC); // Khớp 55.634
        Assert.Equal(19053m, result.TT);  // Khớp 19.053
        Assert.Equal(74687m, result.GT);  // Khớp 74.687
        
        Assert.Equal(46024m, result.TL);  // Khớp 46.024
        Assert.Equal(882824m, result.G);  // Khớp 882.824
        
        Assert.Equal(70626m, result.GTGT); // Khớp 70.626
        Assert.Equal(953450m, result.Gxd); // Khớp 953.450
        
        Assert.Equal(10488m, result.LT);   // Khớp 10.488 (Tuy nhiên có thể chênh lệch 1 đồng do cách làm tròn của phần mềm gốc: G * 1.1% = 882824 * 0.011 = 9711.064???
        // Đợi đã, trong screenshot, Chi phí nhà tạm = 10.488 
        // Tính lại G * 1.1% = 882824 * 1.1% = 9711.064 (Lệch).
        // Hay là Gxd * 1.1% ? 953450 * 1.1% = 10487.95 -> 10488 !!!
        // => À, TT 36 có thể quy định Nhà tạm tính trên Gxd (Sau thuế) hoặc Gx 1.1%?
        // Hãy Kiểm tra kỹ Screenshot 2: "Chi phí nhà tạm để ở và điều hành thi công (G x 1,1...)"
        // Screenshot bị cắt, "G x 1,..." có thể là (Gxd x 1,1%) hoặc (G x 1,...) 
        // 953450 * 1.1% = 10487.95 => Làm tròn = 10488. (Vậy là tính trên Gxd).
    }

    [Fact]
    public void ChiPhiNhaTam_TheoQuyetDinh1538_QD_BXD()
    {
        // Kiểm tra tính tuân thủ theo Quyết định số 1538/QĐ-BXD ngày 28/8/2026:
        // Đính chính công thức tại Bảng 3.8: GXDTT × Tỷ lệ × (1 + TGTGT)
        var calc = new ChiPhiXayDungCalc();
        decimal gxdtt = 1_000_000_000m; // 1 tỷ chi phí XD trước thuế
        decimal tiLeNhaTam = 1.2m;      // 1.2%
        decimal thueSuatGTGT = 10m;     // 10%

        var result = calc.Tinh(
            tongVL: gxdtt,
            tongNC: 0m,
            tongMay: 0m,
            tiLeCPC: 0m,
            tiLeTT: 0m,
            tiLeTNCTTT: 0m,
            tiLeGTGT: thueSuatGTGT,
            tiLeNhaTam: tiLeNhaTam);

        // Công thức theo QĐ 1538: GXDTT × Tỷ lệ × (1 + TGTGT)
        decimal expectedLT = Math.Round(gxdtt * (tiLeNhaTam / 100m) * (1m + thueSuatGTGT / 100m), 0, MidpointRounding.AwayFromZero);
        // = 1.000.000.000 × 1.2% × (1 + 10%) = 12.000.000 × 1.1 = 13.200.000 đ
        Assert.Equal(13_200_000m, expectedLT);
        Assert.Equal(expectedLT, result.LT);
    }
}

