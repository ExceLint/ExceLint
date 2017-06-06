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
            Assert.AreEqual(180, ClusterColorer.Angles(0, 360).Take(1).First());
        }

        [TestMethod]
        public void TestAngleGeneratorMultipleTimes()
        {
            var angles = ClusterColorer.Angles(0, 360);

            Assert.AreEqual(180, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(90, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(270, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(45, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(225, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(135, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(315, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(22.5, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(202.5, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(112.5, angles.Take(1).First());

            angles = angles.Skip(1);

            Assert.AreEqual(292.5, angles.Take(1).First());
        }
    }
}
