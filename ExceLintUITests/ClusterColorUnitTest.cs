using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExceLintUI;
using static ExceLintUI.ColorCalc;
using System.Drawing;

namespace ExceLintUITests
{
    [TestClass]
    public class ClusterColorUnitTest
    {
        private void ColorAssert(Color colorShouldBe, Color colorIs, double delta)
        {
            Assert.AreEqual(colorShouldBe.A, colorIs.A, 0.0);  // transparency should always be the same
            Assert.AreEqual(colorShouldBe.R, colorIs.R, delta);
            Assert.AreEqual(colorShouldBe.G, colorIs.G, delta);
            Assert.AreEqual(colorShouldBe.B, colorIs.B, delta);
        }
        
        [TestMethod]
        public void TestAngleGenerator()
        {
            var angles = new AngleGenerator(0, 360);
            Assert.AreEqual(180, angles.NextAngle());
        }

        [TestMethod]
        public void TestAngleGeneratorSequence()
        {
            var angles = new AngleGenerator(0, 360);

            Assert.AreEqual(180, angles.NextAngle());
            Assert.AreEqual(90, angles.NextAngle());
            Assert.AreEqual(270, angles.NextAngle());
            Assert.AreEqual(45, angles.NextAngle());
            Assert.AreEqual(225, angles.NextAngle());
            Assert.AreEqual(135, angles.NextAngle());
            Assert.AreEqual(315, angles.NextAngle());
            Assert.AreEqual(22.5, angles.NextAngle());
            Assert.AreEqual(202.5, angles.NextAngle());
            Assert.AreEqual(112.5, angles.NextAngle());
            Assert.AreEqual(292.5, angles.NextAngle());
            Assert.AreEqual(67.5, angles.NextAngle());
        }

        [TestMethod]
        public void TestColorSequence()
        {
            double SATURATION = 1.0;
            double LUMINOSITY = 0.5;

            var angles = new AngleGenerator(0, 360);

            ColorAssert(
                RGBtoColor(new RGB(0, 255, 255)),
                HSLtoColor(new HSL(angles.NextAngle(), SATURATION, LUMINOSITY)),
                1.0
            );
            ColorAssert(
                RGBtoColor(new RGB(128, 255, 0)),
                HSLtoColor(new HSL(angles.NextAngle(), SATURATION, LUMINOSITY)),
                1.0
            );
            ColorAssert(
                RGBtoColor(new RGB(128, 0, 255)),
                HSLtoColor(new HSL(angles.NextAngle(), SATURATION, LUMINOSITY)),
                1.0
            );
            ColorAssert(
                RGBtoColor(new RGB(255, 191, 0)),
                HSLtoColor(new HSL(angles.NextAngle(), SATURATION, LUMINOSITY)),
                1.0
            );
        }
    }
}
