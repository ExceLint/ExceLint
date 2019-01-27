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
        [TestMethod]
        public void FormulaRoundTrip()
        {
            
            using (var app = new Application())
            {
                using (var wb = app.OpenWorkbook(@"..\..\TestFiles\OneFormula.xlsx"))
                {
                    var graph = wb.buildDependenceGraph().Worksheets[0];
                    var addr = AST.Address.FromA1String("C5", graph.Worksheet, graph.Workbook, graph.Path);
                    var f = graph.getFormulaAtAddress(addr);
                    Assert.AreEqual("=RAND()", f);
                }
            }
        }

        [TestMethod]
        public void ValueRoundTrip()
        {
            using (var app = new Application())
            {
                using (var wb = app.OpenWorkbook(@"..\..\TestFiles\OneValue.xlsx"))
                {
                    var graph = wb.buildDependenceGraph().Worksheets[0];
                    var addr = AST.Address.FromA1String("F6", graph.Worksheet, graph.Workbook, graph.Path);
                    var v = graph.Values[addr];
                    Assert.AreEqual("7384394", v);
                }
            }
        }
    }
}
