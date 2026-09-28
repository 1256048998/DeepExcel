using System;
using DeepExcel.AddIn.Diagnostics;
using Microsoft.Win32;

namespace DeepExcel.AddIn.Bridge
{
    /// <summary>
    /// 面板跟随 Office 的界面主题（文件 → 账户 → Office 主题）。
    /// 回 "light" / "dark" / "system"；"system" 由面板自己按 prefers-color-scheme 判断
    /// （WebView2 跟随 Windows 的应用主题）。读不到就当 "system"，不猜。
    /// </summary>
    internal static class OfficeTheme
    {
        private const string KeyPath = @"Software\Microsoft\Office\16.0\Common";
        private const string ValueName = "UI Theme";

        /// <summary>
        /// Office 16 的 UI Theme 取值：0 彩色、3 深灰、4 黑色、5 白色、6 使用系统设置。
        /// 1 / 2 是 2013 时代的浅灰 / 深灰，老配置迁移过来可能还留着。
        /// </summary>
        public static string FromRegistryValue(int? value)
        {
            switch (value)
            {
                case 0:
                case 1:
                case 5:
                    return "light";
                case 2:
                case 3:
                case 4:
                    return "dark";
                default:
                    return "system";
            }
        }

        public static string Read()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    var raw = key?.GetValue(ValueName);
                    return FromRegistryValue(raw is int i ? i : (int?)null);
                }
            }
            catch (Exception ex)
            {
                Logger.Instance.Warning("OfficeTheme", $"read UI Theme failed: {ex.Message}");
                return "system";
            }
        }
    }
}
