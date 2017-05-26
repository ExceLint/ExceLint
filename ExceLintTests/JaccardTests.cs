using System;
using ExceLint;
using ExceLintFileFormats;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExceLintTests
{
    [TestClass]
    public class JaccardTests
    {
        [TestMethod]
        public void JaccardIndexIsSane()
        {
            var clustering = Clustering.readClustering(@"..\..\TestData\act3_lab23_poseyxls_clustering_excelint.csv");
            var correspondence = CommonFunctions.JaccardCorrespondence(clustering, clustering);
            Assert.AreEqual(1.0, CommonFunctions.ClusteringJaccardIndex(clustering, clustering, correspondence));
        }
    }
}
