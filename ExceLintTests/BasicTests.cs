using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using COMWrapper;

namespace ExceLintTests
{
    [TestClass]
    public class BasicTests
    {
        [TestMethod]
        [Ignore]    // painfully slow at the moment
        public void canBuildDependenceGraph()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\7_gradebook_xls.xlsx");
            var graph = wb.buildDependenceGraph();
            var dotty = graph.ToDOT();
            Assert.IsTrue(dotty.Length != 0);
        }
    }
}
