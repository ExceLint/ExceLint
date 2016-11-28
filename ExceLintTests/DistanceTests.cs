using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExceLintTests
{
    [TestClass]
    public class DistanceTests
    {
        [TestMethod]
        public void TestEuclideanDistanceForScalars()
        {
            var h1 = Feature.makeNum(1.0);
            var h2 = Feature.makeNum(0.0);
            var dist = ExceLint.ModelBuilder.euclideanDistance(h1, h2);
            Assert.AreEqual(dist, 1.0);
        }

        [TestMethod]
        public void TestEuclideanDistanceForVectors()
        {
            var v1 = Feature.makeVector(0.0, 0.0, 0.0);
            var v2 = Feature.makeVector(1.0, 1.0, 1.0);
            var dist = ExceLint.ModelBuilder.euclideanDistance(v1, v2);
            Assert.AreEqual(dist, Math.Sqrt(3.0));
        }

        [TestMethod]
        public void TestEuclideanDistanceForSpatialVectors()
        {
            var v1 = Feature.makeSpatialVector(0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
            var v2 = Feature.makeSpatialVector(1.0, -1.0, 1.0, -1.0, 1.0, -1.0);
            var dist = ExceLint.ModelBuilder.euclideanDistance(v1, v2);
            Assert.AreEqual(dist, Math.Sqrt(6.0));
        }
    }
}
