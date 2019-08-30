using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.Xlsx
{
    public static class ColorHelper
    {
        public static HlsColor RgbToHls(Color rgbColor)
        {
            var hlsColor = new HlsColor();

            double r = (double)rgbColor.R / 255;
            double g = (double)rgbColor.G / 255;
            double b = (double)rgbColor.B / 255;
            double a = (double)rgbColor.A / 255;

            double min = Math.Min(r, Math.Min(g, b));
            double max = Math.Max(r, Math.Max(g, b));
            double delta = max - min;

            if (max == min)
            {
                hlsColor.H = 0;
                hlsColor.S = 0;
                hlsColor.L = max;

                return hlsColor;
            }

            hlsColor.L = (min + max) / 2;

            if (hlsColor.L < 0.5)
            {
                hlsColor.S = delta / (max + min);
            }

            else
            {
                hlsColor.S = delta / (2.0 - max - min);
            }

            if (r == max) hlsColor.H = (g - b) / delta;
            if (g == max) hlsColor.H = 2.0 + (b - r) / delta;
            if (b == max) hlsColor.H = 4.0 + (r - g) / delta;

            hlsColor.H *= 60;

            if (hlsColor.H < 0) hlsColor.H += 360;

            hlsColor.A = a;

            return hlsColor;
        }

        public static Color HlsToRgb(HlsColor hlsColor)
        {
            var rgbColor = new Color();

            if (hlsColor.S == 0)
            {
                rgbColor = Color.FromArgb((int)(hlsColor.A * 255), (int)(hlsColor.L * 255), (int)(hlsColor.L * 255),
                (int)(hlsColor.L * 255));
                return rgbColor;
            }

            double t1;
            if (hlsColor.L < 0.5)
            {
                t1 = hlsColor.L * (1.0 + hlsColor.S);
            }
            else
            {
                t1 = hlsColor.L + hlsColor.S - (hlsColor.L * hlsColor.S);
            }

            double t2 = 2.0 * hlsColor.L - t1;
            double h = hlsColor.H / 360;
            double tR = h + (1.0 / 3.0);
            double r = SetColor(t1, t2, tR);
            double tG = h;
            double g = SetColor(t1, t2, tG);
            double tB = h - (1.0 / 3.0);
            double b = SetColor(t1, t2, tB);

            rgbColor = Color.FromArgb((int)(hlsColor.A * 255), (int)(r * 255), (int)(g * 255), (int)(b * 255));
            return rgbColor;
        }

        public static Color ApplyTintToRgb(Color color, double tint)
        {
            HlsColor rgbToHls = RgbToHls(color);
            rgbToHls.L = ApplyTintToLum(tint, (float)rgbToHls.L * 255) / 255;
            Color hlsToRgb = HlsToRgb(rgbToHls);
            return hlsToRgb;
        }

        private static double SetColor(double t1, double t2, double t3)
        {
            if (t3 < 0) t3 += 1.0;
            if (t3 > 1) t3 -= 1.0;

            double color;

            if (6.0 * t3 < 1)
            {
                color = t2 + (t1 - t2) * 6.0 * t3;
            }
            else if (2.0 * t3 < 1)
            {
                color = t1;
            }
            else if (3.0 * t3 < 2)
            {
                color = t2 + (t1 - t2) * ((2.0 / 3.0) - t3) * 6.0;
            }
            else
            {
                color = t2;
            }

            return color;
        }

        private static double ApplyTintToLum(double tint, float lum)
        {
            double lum1 = 0;

            if (tint < 0)
            {
                lum1 = lum * (1.0 + tint);
            }
            else
            {
                lum1 = lum * (1.0 - tint) + (255 - 255 * (1.0 - tint));
            }

            return lum1;
        }
    }
}
