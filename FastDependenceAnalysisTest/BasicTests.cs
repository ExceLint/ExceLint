using System;
using System.Runtime.Remoting;
using COMWrapper;
using FastDependenceAnalysis;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FastDependenceAnalysisTest
{
    [TestClass]
    public class BasicTests
    {
        public Graph FormulaGraph()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestFiles\OneFormula.xlsx");
            return wb.buildDependenceGraph().Worksheets[0];
        }

        public Graph ValueGraph()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestFiles\OneValue.xlsx");
            return wb.buildDependenceGraph().Worksheets[0];
        }

        [TestMethod]
        public void FormulaRoundTrip()
        {
            var graph = FormulaGraph();
            var addr = AST.Address.FromA1String("C5", graph.Worksheet, graph.Workbook, graph.Path);
            var f = graph.getFormulaAtAddress(addr);
            Assert.AreEqual("=RAND()", f);
        }
    }
}
