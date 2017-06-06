using System;
using System.Drawing;
using ExceLintUI;
using static ExceLintUI.ColorCalc;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExceLintUITests
{
    [TestClass]
    public class ColorUnitTest
    {
        [TestMethod]
        public void ColorToRGBTest()
        {
            var red = Color.Red;
            var red_rgb = ColorToRGB(red);

            Assert.AreEqual(255, red_rgb.Red);
            Assert.AreEqual(0, red_rgb.Green);
            Assert.AreEqual(0, red_rgb.Blue);
        }

        [TestMethod]
        public void RGBToColorTest()
        {
            var red_rgb = new RGB(255, 0, 0);
            var red = RGBtoColor(red_rgb);
            var real_red = Color.Red;

            Assert.AreEqual(real_red.A, red.A);
            Assert.AreEqual(real_red.R, red.R);
            Assert.AreEqual(real_red.G, red.G);
            Assert.AreEqual(real_red.B, red.B);
        }

        [TestMethod]
        public void RGBtoHSLTest()
        {
            var c_rgb = new RGB(24, 98, 118);
            var c_hsl = RGBtoHSL(c_rgb);

            Assert.AreEqual(193, c_hsl.Hue, 0.5);
            Assert.AreEqual(0.672, c_hsl.Saturation, 0.05);
            Assert.AreEqual(0.275, c_hsl.Luminosity, 0.05);
        }

        [TestMethod]
        public void HSLtoRGBTest()
        {
            var c_hsl = new HSL(193, 0.67, 0.28);
            var c_rgb = HSLtoRGB(c_hsl);

            Assert.AreEqual(24, c_rgb.Red);
            Assert.AreEqual(98, c_rgb.Green);
            Assert.AreEqual(119, c_rgb.Blue);
        }

        [TestMethod]
        public void ComplementTest()
        {
            var c_rgb = new RGB(24, 98, 118);
            var c_hsl = RGBtoHSL(c_rgb);

            var comp_rgb = GetComplementaryColor(c_rgb);
            var comp_hsl = RGBtoHSL(comp_rgb);

            Assert.AreEqual(c_hsl.Hue, (comp_hsl.Hue + 180) % 360, 0.5);
            Assert.AreEqual(c_hsl.Saturation, comp_hsl.Saturation, 0.05);
            Assert.AreEqual(c_hsl.Luminosity, comp_hsl.Luminosity, 0.05);
        }

        [TestMethod]
        public void ColorComplementTest()
        {
            var c = Color.Goldenrod;
            var c_comp = GetComplementaryColor(c);

            Assert.AreEqual(32, c_comp.R);
            Assert.AreEqual(84, c_comp.G);
            Assert.AreEqual(218, c_comp.B);
        }
    }
}
