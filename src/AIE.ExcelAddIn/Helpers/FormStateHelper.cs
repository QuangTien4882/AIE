using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace AIE.ExcelAddIn.Helpers
{
    public class FormBoundInfo
    {
        public int Width { get; set; }
        public int Height { get; set; }
        public bool IsMaximized { get; set; }
    }

    /// <summary>
    /// Helper tự động ghi nhớ và khôi phục kích thước / trạng thái cửa sổ các Modal
    /// Dữ liệu được lưu tại %LOCALAPPDATA%\AIE_DuToan\form_bounds.json
    /// </summary>
    public static class FormStateHelper
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIE_DuToan");
        private static readonly string SettingsFile = Path.Combine(SettingsDir, "form_bounds.json");
        private static readonly object FileLock = new object();

        private static Dictionary<string, FormBoundInfo> LoadAllBounds()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile);
                    var data = JsonConvert.DeserializeObject<Dictionary<string, FormBoundInfo>>(json);
                    if (data != null) return data;
                }
            }
            catch { }
            return new Dictionary<string, FormBoundInfo>(StringComparer.OrdinalIgnoreCase);
        }

        private static void SaveBound(string key, FormBoundInfo info)
        {
            try
            {
                lock (FileLock)
                {
                    if (!Directory.Exists(SettingsDir)) Directory.CreateDirectory(SettingsDir);
                    var data = LoadAllBounds();
                    data[key] = info;
                    File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(data, Formatting.Indented));
                }
            }
            catch { }
        }

        /// <summary>
        /// Gắn cơ chế tự động ghi nhớ và phục hồi kích thước cho một Form WinForms
        /// </summary>
        public static void Attach(Form form, string formKey = null)
        {
            if (form == null) return;
            string key = string.IsNullOrEmpty(formKey) ? form.GetType().Name : formKey;

            // 1. Phục hồi kích thước khi Form mở
            form.Load += (s, e) =>
            {
                try
                {
                    var data = LoadAllBounds();
                    if (data.TryGetValue(key, out var info) && info != null)
                    {
                        var area = Screen.FromControl(form).WorkingArea;

                        if (info.IsMaximized && form.MaximizeBox)
                        {
                            form.WindowState = FormWindowState.Maximized;
                        }
                        else if (info.Width > 0 && info.Height > 0)
                        {
                            int w = Math.Max(form.MinimumSize.Width, info.Width);
                            int h = Math.Max(form.MinimumSize.Height, info.Height);

                            // Đảm bảo không tràn ra ngoài màn hình làm việc
                            w = Math.Min(w, area.Width - 20);
                            h = Math.Min(h, area.Height - 40);

                            form.Size = new Size(w, h);
                            form.StartPosition = FormStartPosition.CenterScreen;
                        }
                    }
                }
                catch { }
            };

            // 2. Lưu kích thước khi Form đóng
            form.FormClosing += (s, e) =>
            {
                try
                {
                    var info = new FormBoundInfo
                    {
                        IsMaximized = form.WindowState == FormWindowState.Maximized
                    };

                    if (form.WindowState == FormWindowState.Normal)
                    {
                        info.Width = form.Width;
                        info.Height = form.Height;
                    }
                    else if (form.WindowState == FormWindowState.Maximized && form.RestoreBounds.Width > 0)
                    {
                        info.Width = form.RestoreBounds.Width;
                        info.Height = form.RestoreBounds.Height;
                    }

                    if (info.Width > 0 && info.Height > 0)
                    {
                        SaveBound(key, info);
                    }
                }
                catch { }
            };
        }
    }
}
