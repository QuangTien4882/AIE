using Xunit;
using AIE.Data.ImportExport;
using System.IO;
using System;

namespace AIE.Data.Tests;

public class ExcelTemplateGeneratorTests
{
    [Fact]
    public void Can_Generate_All_Templates()
    {
        var generator = new ExcelTemplateGenerator();
        var outDir = Path.Combine(Path.GetTempPath(), "AIE_Templates_" + Guid.NewGuid());
        
        generator.GenerateAllTemplates(outDir);

        Assert.True(File.Exists(Path.Combine(outDir, "1_DinhMucCongTac.xlsx")));
        Assert.True(File.Exists(Path.Combine(outDir, "2_GiaVatLieu.xlsx")));
        Assert.True(File.Exists(Path.Combine(outDir, "3_GiaNhanCong.xlsx")));
        Assert.True(File.Exists(Path.Combine(outDir, "4_GiaMayThiCong.xlsx")));
        Assert.True(File.Exists(Path.Combine(outDir, "5_TiLePhanTram.xlsx")));
        Assert.True(File.Exists(Path.Combine(outDir, "6_DinhMucTuVan.xlsx")));

        // Cleanup
        Directory.Delete(outDir, true);
    }
}
