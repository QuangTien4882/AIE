using System;
using System.Collections.Generic;
using System.IO;
using AIE.Core.Models;
using Newtonsoft.Json;

namespace AIE.ExcelAddIn.Services;

public class VanChuyenOToSavedConfig
{
    public string MaDinhMuc { get; set; } = string.Empty;
    public string MaMay { get; set; } = string.Empty;
    public decimal DonGiaCaMay { get; set; }
    public decimal TongCuLyKm { get; set; }
    public List<CungDuongVanChuyen> CungDuongs { get; set; } = new();
    public decimal CuocVCOTo { get; set; }
}

public class VanChuyenBoSavedConfig
{
    public string MaDinhMuc { get; set; } = string.Empty;
    public decimal TongCuLyMet { get; set; }
    public decimal CuLyMet { get; set; }
    public decimal HeSoDiaHinh { get; set; } = 1.0m;
    public int SoTang { get; set; } = 1;
    public List<DoanVanChuyenBo> DoanBos { get; set; } = new();
    public decimal CuocVCBo { get; set; }
    public decimal DonGiaNhanCong { get; set; }
    public string MaNhanCong { get; set; } = string.Empty;
}

public class TuyenDuongTemplate
{
    public string TenTemplate { get; set; } = string.Empty;
    public decimal TongCuLyKm { get; set; }
    public List<CungDuongVanChuyen> CungDuongs { get; set; } = new();
    public DateTime NgayTao { get; set; } = DateTime.Now;

    public override string ToString() => $"{TenTemplate} ({TongCuLyKm:0.###} km)";
}

public static class VanChuyenStorage
{
    private static Dictionary<string, VanChuyenOToSavedConfig>? _cacheOTo;
    private static Dictionary<string, VanChuyenBoSavedConfig>? _cacheBo;
    private static List<TuyenDuongTemplate>? _templates;
    private static readonly object _lock = new();

    private static string GetDir()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        return dir;
    }

    private static string NormalizeKey(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        return name.Trim().TrimStart('-').Trim().ToLowerInvariant();
    }

    // ================= Ô TÔ =================
    private static void EnsureLoadedOTo()
    {
        if (_cacheOTo != null) return;
        lock (_lock)
        {
            if (_cacheOTo != null) return;
            _cacheOTo = new Dictionary<string, VanChuyenOToSavedConfig>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var file = Path.Combine(GetDir(), "van_chuyen_oto_config.json");
                if (File.Exists(file))
                {
                    var json = File.ReadAllText(file);
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, VanChuyenOToSavedConfig>>(json);
                    if (dict != null)
                    {
                        foreach (var kv in dict) _cacheOTo[kv.Key] = kv.Value;
                    }
                }
            }
            catch { }
        }
    }

    public static VanChuyenOToSavedConfig? GetConfigOTo(string tenVatTu)
    {
        EnsureLoadedOTo();
        var key = NormalizeKey(tenVatTu);
        if (string.IsNullOrEmpty(key)) return null;
        lock (_lock)
        {
            return _cacheOTo != null && _cacheOTo.TryGetValue(key, out var cfg) ? cfg : null;
        }
    }

    public static void SaveConfigOTo(string tenVatTu, string maDinhMuc, string maMay, decimal donGiaCaMay, List<CungDuongVanChuyen> cungDuongs, decimal cuocVCOTo)
    {
        if (string.IsNullOrWhiteSpace(tenVatTu)) return;
        EnsureLoadedOTo();
        var key = NormalizeKey(tenVatTu);
        lock (_lock)
        {
            _cacheOTo ??= new Dictionary<string, VanChuyenOToSavedConfig>(StringComparer.OrdinalIgnoreCase);
            _cacheOTo[key] = new VanChuyenOToSavedConfig
            {
                MaDinhMuc = maDinhMuc,
                MaMay = maMay,
                DonGiaCaMay = donGiaCaMay,
                TongCuLyKm = cungDuongs != null ? cungDuongs.Sum(c => c.CuLyKm) : 0,
                CungDuongs = cungDuongs ?? new List<CungDuongVanChuyen>(),
                CuocVCOTo = cuocVCOTo
            };

            try
            {
                var file = Path.Combine(GetDir(), "van_chuyen_oto_config.json");
                var json = JsonConvert.SerializeObject(_cacheOTo, Formatting.Indented);
                File.WriteAllText(file, json);
            }
            catch { }
        }
    }

    // ================= MẪU TUYẾN ĐƯỜNG (TEMPLATE) =================
    private static void EnsureLoadedTemplates()
    {
        if (_templates != null) return;
        lock (_lock)
        {
            if (_templates != null) return;
            _templates = new List<TuyenDuongTemplate>();
            try
            {
                var file = Path.Combine(GetDir(), "tuyen_duong_templates.json");
                if (File.Exists(file))
                {
                    var json = File.ReadAllText(file);
                    var list = JsonConvert.DeserializeObject<List<TuyenDuongTemplate>>(json);
                    if (list != null) _templates = list;
                }
            }
            catch { }
        }
    }

    public static List<TuyenDuongTemplate> GetAllTemplates()
    {
        EnsureLoadedTemplates();
        lock (_lock)
        {
            return _templates != null ? new List<TuyenDuongTemplate>(_templates) : new List<TuyenDuongTemplate>();
        }
    }

    public static void SaveTemplate(string tenTemplate, List<CungDuongVanChuyen> cungDuongs)
    {
        if (string.IsNullOrWhiteSpace(tenTemplate) || cungDuongs == null) return;
        EnsureLoadedTemplates();
        lock (_lock)
        {
            _templates ??= new List<TuyenDuongTemplate>();
            var existing = _templates.Find(t => t.TenTemplate.Equals(tenTemplate.Trim(), StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                existing.CungDuongs = new List<CungDuongVanChuyen>(cungDuongs);
                existing.TongCuLyKm = cungDuongs.Sum(c => c.CuLyKm);
                existing.NgayTao = DateTime.Now;
            }
            else
            {
                _templates.Add(new TuyenDuongTemplate
                {
                    TenTemplate = tenTemplate.Trim(),
                    TongCuLyKm = cungDuongs.Sum(c => c.CuLyKm),
                    CungDuongs = new List<CungDuongVanChuyen>(cungDuongs),
                    NgayTao = DateTime.Now
                });
            }

            try
            {
                var file = Path.Combine(GetDir(), "tuyen_duong_templates.json");
                var json = JsonConvert.SerializeObject(_templates, Formatting.Indented);
                File.WriteAllText(file, json);
            }
            catch { }
        }
    }

    public static void DeleteTemplate(string tenTemplate)
    {
        if (string.IsNullOrWhiteSpace(tenTemplate)) return;
        EnsureLoadedTemplates();
        lock (_lock)
        {
            if (_templates == null) return;
            _templates.RemoveAll(t => t.TenTemplate.Equals(tenTemplate.Trim(), StringComparison.OrdinalIgnoreCase));
            try
            {
                var file = Path.Combine(GetDir(), "tuyen_duong_templates.json");
                var json = JsonConvert.SerializeObject(_templates, Formatting.Indented);
                File.WriteAllText(file, json);
            }
            catch { }
        }
    }

    // ================= BỘ =================
    private static void EnsureLoadedBo()
    {
        if (_cacheBo != null) return;
        lock (_lock)
        {
            if (_cacheBo != null) return;
            _cacheBo = new Dictionary<string, VanChuyenBoSavedConfig>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var file = Path.Combine(GetDir(), "van_chuyen_bo_config.json");
                if (File.Exists(file))
                {
                    var json = File.ReadAllText(file);
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, VanChuyenBoSavedConfig>>(json);
                    if (dict != null)
                    {
                        foreach (var kv in dict) _cacheBo[kv.Key] = kv.Value;
                    }
                }
            }
            catch { }
        }
    }

    public static VanChuyenBoSavedConfig? GetConfigBo(string tenVatTu)
    {
        EnsureLoadedBo();
        var key = NormalizeKey(tenVatTu);
        if (string.IsNullOrEmpty(key)) return null;
        lock (_lock)
        {
            return _cacheBo != null && _cacheBo.TryGetValue(key, out var cfg) ? cfg : null;
        }
    }

    public static void SaveConfigBo(
        string tenVatTu, 
        string maDinhMuc, 
        decimal cuLyMet, 
        decimal heSoDiaHinh, 
        int soTang, 
        decimal cuocVCBo,
        decimal donGiaNhanCong = 0,
        string maNhanCong = "")
    {
        if (string.IsNullOrWhiteSpace(tenVatTu)) return;
        EnsureLoadedBo();
        var key = NormalizeKey(tenVatTu);
        lock (_lock)
        {
            _cacheBo ??= new Dictionary<string, VanChuyenBoSavedConfig>(StringComparer.OrdinalIgnoreCase);
            var existing = _cacheBo.TryGetValue(key, out var old) ? old : null;
            _cacheBo[key] = new VanChuyenBoSavedConfig
            {
                MaDinhMuc = maDinhMuc,
                CuLyMet = cuLyMet,
                HeSoDiaHinh = heSoDiaHinh,
                SoTang = soTang,
                CuocVCBo = cuocVCBo,
                DonGiaNhanCong = donGiaNhanCong > 0 ? donGiaNhanCong : (existing?.DonGiaNhanCong ?? 0),
                MaNhanCong = !string.IsNullOrEmpty(maNhanCong) ? maNhanCong : (existing?.MaNhanCong ?? "")
            };

            try
            {
                var file = Path.Combine(GetDir(), "van_chuyen_bo_config.json");
                var json = JsonConvert.SerializeObject(_cacheBo, Formatting.Indented);
                File.WriteAllText(file, json);
            }
            catch { }
        }
    }
}
