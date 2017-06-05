using System;
using System.Linq;
using System.Drawing;
using static ExceLintUI.RibbonHelper;

namespace ExceLintUI
{
    struct HSL
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

    struct RGB
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

    static class ColorCalc
    {
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
                hue = (hue * 60) % 360;
            }

            return new HSL(hue, saturation, luminance);
        }

        public static RGB HSLtoRGB(HSL hsl)
        {
            // derived from: https://en.wikipedia.org/wiki/HSL_and_HSV#From_HSV
            double chroma = hsl.Saturation * hsl.Luminosity;
            double huePrime = hsl.Hue / 60;
            double X = chroma * (1 - Math.Abs(huePrime % 2 - 1));
            double R1 = 0, G1 = 0, B1 = 0;

            if (huePrime >= 0 && huePrime < 1)
            {
                R1 = chroma;
                G1 = X;
                B1 = 0;
            }
            else if (huePrime >= 1 && huePrime < 2)
            {
                R1 = X;
                G1 = chroma;
                B1 = 0;
            }
            else if (huePrime >= 2 && huePrime < 3)
            {
                R1 = 0;
                G1 = chroma;
                B1 = X;
            }
            else if (huePrime >= 3 && huePrime < 4)
            {
                R1 = 0;
                G1 = X;
                B1 = chroma;
            }
            else if (huePrime >= 4 && huePrime < 5)
            {
                R1 = X;
                G1 = 0;
                B1 = chroma;
            }
            else
            {
                R1 = chroma;
                G1 = 0;
                B1 = X;
            }

            double m = hsl.Luminosity - chroma;

            double R = R1 + m,
                   G = G1 + m,
                   B = B1 + m;

            System.Diagnostics.Debug.Assert(R >= 0 && R <= 1);
            System.Diagnostics.Debug.Assert(G >= 0 && G <= 1);
            System.Diagnostics.Debug.Assert(B >= 0 && B <= 1);

            byte r = Convert.ToByte(R * 255),
                 g = Convert.ToByte(G * 255),
                 b = Convert.ToByte(B * 255);

            return new RGB(r, g, b);
        }

        public static RGB GetComplementaryColor(RGB rgb)
        {
            // convert to HSL
            HSL hsl = RGBtoHSL(rgb);

            // find complementary color
            HSL hslc = new HSL(hsl.Hue, (hsl.Saturation - 180.0) % 360.0, hsl.Luminosity);

            // convert back to RGB and return
            return HSLtoRGB(hsl);
        }

        public static Color GetComplementaryColor(Color c)
        {
            // convert to RGB
            int argbi = c.ToArgb();
            byte[] argb_bs = BitConverter.GetBytes(argbi);
            byte red = BitConverter.IsLittleEndian ? argb_bs[2] : argb_bs[1];
            byte green = BitConverter.IsLittleEndian ? argb_bs[1] : argb_bs[2];
            byte blue = BitConverter.IsLittleEndian ? argb_bs[0] : argb_bs[3];
            RGB rgb = new RGB(red, green, blue);

            // find complementary color
            RGB rgbc = GetComplementaryColor(rgb);

            // convert back to Color
            return Color.FromArgb(255, rgb.Red, rgb.Green, rgb.Blue);
        }
    }
}
