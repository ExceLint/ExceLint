using System;
using COMWrapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Resultant = ExceLint.Countable.CVectorResultant;

namespace ExceLintTests
{
    [TestClass]
    public class CVectorTests
    {
        private Depends.DAG SimpleDAG()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\SimpleWorkbook.xlsx");
            var graph = wb.buildDependenceGraph();
            return graph;
        }

        private Depends.DAG SimpleDAGWithConstant()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\SimpleWorkbookWithConstant.xlsx");
            var graph = wb.buildDependenceGraph();
            return graph;
        }

        [TestMethod]
        public void BasicCVectorResultantTest()
        {
            // tests that A1 in cell B1 on the same sheet returns the relative resultant vector (-1,1,0,0)
            var dag = SimpleDAG();
            var wbname = dag.getWorkbookName();
            var wsname = dag.getWorksheetNames()[0];
            var path = dag.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(1, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var resultant = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula, dag);
            var resultant_shouldbe = Resultant.NewCVectorResultant(-1,0,0,0);
            Assert.AreEqual(resultant_shouldbe, resultant);
        }

        [TestMethod]
        public void BasicCVectorResultantTest2()
        {
            // tests that A1 in cell B1 on the same sheet returns the relative resultant vector (-1,1,0,1)
            var dag = SimpleDAGWithConstant();
            var wbname = dag.getWorkbookName();
            var wsname = dag.getWorksheetNames()[0];
            var path = dag.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(1, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var resultant = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula, dag);
            var resultant_shouldbe = Resultant.NewCVectorResultant(-1, 0, 0, 1);
            Assert.AreEqual(resultant_shouldbe, resultant);
        }
    }
}
