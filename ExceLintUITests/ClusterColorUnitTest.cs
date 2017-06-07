using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExceLintUI;
using System.Linq;

namespace ExceLintUITests
{
    [TestClass]
    public class ClusterColorUnitTest
    {

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
    }
}
