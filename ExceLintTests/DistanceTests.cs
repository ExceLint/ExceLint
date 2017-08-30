using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ExceLint;

namespace ExceLintTests
{
    [TestClass]
    public class DistanceTests
    {
        [TestMethod]
        public void TestEuclideanDistanceForScalars()
        {
            var h1 = FeatureUtil.makeNum(1.0);
            var h2 = FeatureUtil.makeNum(0.0);
            var dist = SpectralModelBuilder.euclideanDistance(h1, h2);
            Assert.AreEqual(dist, 1.0);
        }

        [TestMethod]
        public void TestEuclideanDistanceForVectors()
        {
            var v1 = FeatureUtil.makeVector(0.0, 0.0, 0.0);
            var v2 = FeatureUtil.makeVector(1.0, 1.0, 1.0);
            var dist = SpectralModelBuilder.euclideanDistance(v1, v2);
            Assert.AreEqual(dist, Math.Sqrt(3.0));
        }

        [TestMethod]
        public void TestEuclideanDistanceForSpatialVectors()
        {
            var v1 = FeatureUtil.makeSpatialVector(0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
            var v2 = FeatureUtil.makeSpatialVector(1.0, -1.0, 1.0, -1.0, 1.0, -1.0);
            var dist = SpectralModelBuilder.euclideanDistance(v1, v2);
            Assert.AreEqual(dist, Math.Sqrt(6.0));
        }
    }
}
