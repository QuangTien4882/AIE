using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace AIE.ExcelAddIn.Services;

public class BocXepSavedConfig
{
    public string MaDinhMuc { get; set; } = string.Empty;
    public int PhamVi { get; set; }
    public decimal DmNC { get; set; }
    public decimal DmMay { get; set; }
    public string? MaMay { get; set; }
    public decimal ChiPhiBocXep { get; set; }
}

public static class BocXepStorage
{
    private static Dictionary<string, BocXepSavedConfig>? _cache;
    private static readonly object _lock = new();

    private static string GetFilePath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
            "AIE_DuToan");
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        return Path.Combine(dir, "boc_xep_config.json");
    }

    private static string NormalizeKey(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        return name.Trim().TrimStart('-').Trim().ToLowerInvariant();
    }

    private static void EnsureLoaded()
    {
        if (_cache != null) return;
        lock (_lock)
        {
            if (_cache != null) return;
            _cache = new Dictionary<string, BocXepSavedConfig>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var file = GetFilePath();
                if (File.Exists(file))
                {
                    var json = File.ReadAllText(file);
                    var dict = JsonConvert.DeserializeObject<Dictionary<string, BocXepSavedConfig>>(json);
                    if (dict != null)
                    {
                        foreach (var kv in dict)
                            _cache[kv.Key] = kv.Value;
                    }
                }
            }
            catch { }
        }
    }

    public static BocXepSavedConfig? GetConfig(string tenVatTu)
    {
        EnsureLoaded();
        var key = NormalizeKey(tenVatTu);
        if (string.IsNullOrEmpty(key)) return null;
        lock (_lock)
        {
            return _cache != null && _cache.TryGetValue(key, out var cfg) ? cfg : null;
        }
    }

    public static void SaveConfig(string tenVatTu, string? maDinhMuc, int phamVi, decimal dmNC, decimal dmMay, string? maMay, decimal chiPhi)
    {
        if (string.IsNullOrWhiteSpace(tenVatTu) || string.IsNullOrEmpty(maDinhMuc)) return;
        EnsureLoaded();
        var key = NormalizeKey(tenVatTu);
        lock (_lock)
        {
            _cache ??= new Dictionary<string, BocXepSavedConfig>(StringComparer.OrdinalIgnoreCase);
            _cache[key] = new BocXepSavedConfig
            {
                MaDinhMuc = maDinhMuc,
                PhamVi = phamVi,
                DmNC = dmNC,
                DmMay = dmMay,
                MaMay = maMay,
                ChiPhiBocXep = chiPhi
            };

            try
            {
                var file = GetFilePath();
                var json = JsonConvert.SerializeObject(_cache, Formatting.Indented);
                File.WriteAllText(file, json);
            }
            catch { }
        }
    }
}
