using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;

namespace PSWikiTable
{
    internal static class ColorTheme
    {
        // This is the default color theme in Excel. More themes and custom themes may be supported in the future.
        private static readonly Dictionary<eThemeSchemeColor, string> colorTheme = new Dictionary<eThemeSchemeColor, string>()
        {
            { eThemeSchemeColor.Accent1, "#5B9BD5" },
            { eThemeSchemeColor.Accent2, "#ED7D31" },
            { eThemeSchemeColor.Accent3, "#A5A5A5" },
            { eThemeSchemeColor.Accent4, "#FFC000" },
            { eThemeSchemeColor.Accent5, "#4472C4" },
            { eThemeSchemeColor.Accent6, "#70AD47" },
            { eThemeSchemeColor.Background1, "#FFFFFF" },
            { eThemeSchemeColor.Background2, "#E7E6E6" },
            { eThemeSchemeColor.Text1, "#000000" },
            { eThemeSchemeColor.Text2, "#44546A" }
        };

        public static string GetThemeRgb(eThemeSchemeColor themeColor, decimal tint)
        {
            if (colorTheme.TryGetValue(themeColor, out string rgb))
            {
                if (tint > 0.001M || tint < -0.001M)
                {
                    rgb = TintRgb(rgb, tint);
                }
                return rgb;
            }
            return null;
        }

        // This implementation of "tinting" is not a perfect replication of how Excel does it,
        // but it's close enough. Especially darkening (where tint has a minus value) differs
        // from Excel.
        // It would also be more efficient to always work with RGB as separate integer values instead
        // of translating from and to hex string, but EPPlus uses hex string representation and
        // also something about premature optimization...
        private static string TintRgb(string rgb, decimal tint)
        {
            int t = 255;
            if (tint < 0)
            {
                t = 0;
                tint = -tint;
            }

            int r = Convert.ToInt32(rgb.Substring(1, 2), 16);
            int g = Convert.ToInt32(rgb.Substring(3, 2), 16);
            int b = Convert.ToInt32(rgb.Substring(5, 2), 16);

            r = (int)Math.Round((t - r) * tint) + r;
            g = (int)Math.Round((t - g) * tint) + g;
            b = (int)Math.Round((t - b) * tint) + b;

            return $"#{r:X}{g:X}{b:X}";
        }
    }
}
