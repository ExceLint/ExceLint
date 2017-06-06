using System;
using System.Linq;
using System.Drawing;
using static ExceLintUI.RibbonHelper;

namespace ExceLintUI
{
    public struct HSL
    {
        // real number; degrees from 0 to <360
        public double Hue;
        // real number from 0 to 1
        public double Saturation, Luminosity;

        public HSL(double hue, double saturation, double luminosity)
        {
            Hue = hue;
            Saturation = saturation;
            Luminosity = luminosity;
        }
    }

    public struct RGB
    {
        // int from 0 to 255
        public byte Red, Green, Blue;

        public RGB(byte red, byte green, byte blue)
        {
            Red = red;
            Green = green;
            Blue = blue;
        }
    }

    public static class ColorCalc
    {
        // C#'s % operator is remainder, not modulus
        // https://stackoverflow.com/questions/1082917/mod-of-negative-number-is-melting-my-brain/6400477#6400477
        private static double mod(double a, double b)
        {
            return a - b * Math.Floor(a / b);
        }

        public static RGB ColorToRGB(Color c)
        {
            int argbi = c.ToArgb();
            byte[] argb_bs = BitConverter.GetBytes(argbi);
            byte red = BitConverter.IsLittleEndian ? argb_bs[2] : argb_bs[1];
            byte green = BitConverter.IsLittleEndian ? argb_bs[1] : argb_bs[2];
            byte blue = BitConverter.IsLittleEndian ? argb_bs[0] : argb_bs[3];
            return new RGB(red, green, blue);
        }

        public static Color RGBtoColor(RGB rgb)
        {
            return Color.FromArgb(255, rgb.Red, rgb.Green, rgb.Blue);
        }

        public static HSL ColorToHSL(Color c)
        {
            return RGBtoHSL(ColorToRGB(c));
        }

        public static HSL RGBtoHSL(RGB rgb)
        {
            // derived from: http://www.niwa.nu/2013/05/math-behind-colorspace-conversions-rgb-hsl/
            double r_rel = rgb.Red / 255.0,
                   g_rel = rgb.Green / 255.0,
                   b_rel = rgb.Blue / 255.0;

            double[] rgb_rel_arr = { r_rel, g_rel, b_rel };
            double min = rgb_rel_arr.Min();
            double max = rgb_rel_arr.Max();
            int maxidx = ArgMax(rgb_rel_arr);

            double luminance = (max + min) / 2.0;

            // if max == min then hue and saturation are 0
            double saturation = 0,
                   hue = 0;

            if (max != min)
            {
                // compute saturation
                saturation = luminance < 0.5 ? (max - min) / (max + min) : (max - min) / (2.0 - max - min);

                // compute hue
                if (maxidx == 0)
                {
                    // red is max
                    hue = (g_rel - b_rel) / (max - min);
                } else if (maxidx == 1)
                {
                    // green is max
                    hue = 2.0 + (b_rel - r_rel) / (max - min);
                } else
                {
                    // blue is max
                    hue = 4.0 + (r_rel - g_rel) / (max - min);
                }

                // multiply by 60 to make this a proper angle in degrees
                hue = mod((hue * 60), 360);
            }

            return new HSL(hue, saturation, luminance);
        }

        private static double ConvertChannel(double C, double temp1, double temp2)
        {
            if (6 * C < 1.0)
            {
                return temp2 + (temp1 - temp2) * 6 * C;
            }
            else if (2 * C < 1.0)
            {
                return temp1;
            }
            else if (3 * C < 2.0)
            {
                return temp2 + (temp1 - temp2) * (0.666 - C) * 6.0;
            }
            else
            {
                return temp2;
            }
        }

        public static RGB HSLtoRGB(HSL hsl)
        {
            var temp1 = hsl.Luminosity < 0.5 ?
                        hsl.Luminosity * (1.0 + hsl.Saturation) :
                        hsl.Luminosity + hsl.Saturation - hsl.Luminosity * hsl.Saturation;
            var temp2 = 2.0 * hsl.Luminosity - temp1;
            var hue_rel = hsl.Hue / 360;

            var R = mod((hue_rel + 0.333), 1.0);
            var G = hue_rel;
            var B = mod((hue_rel - 0.333), 1.0);

            R = ConvertChannel(R, temp1, temp2);
            G = ConvertChannel(G, temp1, temp2);
            B = ConvertChannel(B, temp1, temp2);

            return new RGB(
                Convert.ToByte(R * 255),
                Convert.ToByte(G * 255),
                Convert.ToByte(B * 255));
        }

        public static RGB GetComplementaryColor(RGB rgb)
        {
            // convert to HSL
            HSL hsl = RGBtoHSL(rgb);

            // find complementary color
            HSL hslc = new HSL(mod((hsl.Hue - 180.0), 360.0), hsl.Saturation, hsl.Luminosity);

            // convert back to RGB and return
            return HSLtoRGB(hslc);
        }

        public static Color GetComplementaryColor(Color c)
        {
            // convert to RGB
            RGB rgb = ColorToRGB(c);

            // find complementary color
            RGB rgbc = GetComplementaryColor(rgb);

            // convert complement back to Color
            return RGBtoColor(rgbc);
        }
    }
}
