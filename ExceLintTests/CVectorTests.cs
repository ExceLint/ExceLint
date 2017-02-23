using System;
using COMWrapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Resultant = ExceLint.Countable.CVectorResultant;
using System.Linq;

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

        private Depends.DAG SimpleDAGWithConstants()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\SimpleWorkbookWithConstants.xlsx");
            var graph = wb.buildDependenceGraph();
            return graph;
        }

        private Depends.DAG DAGWithMultipleFormulasAndConstants()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\SimpleWorkbookWithMultipleFormulasAndConstants.xlsx");
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

        [TestMethod]
        public void BasicCVectorResultantTest3()
        {
            // tests that A1 in cell B1 on the same sheet returns the relative resultant vector (-1,1,0,3)
            var dag = SimpleDAGWithConstants();
            var wbname = dag.getWorkbookName();
            var wsname = dag.getWorksheetNames()[0];
            var path = dag.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(1, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var resultant = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula, dag);
            var resultant_shouldbe = Resultant.NewCVectorResultant(-1, 0, 0, 3);
            Assert.AreEqual(resultant_shouldbe, resultant);
        }

        [TestMethod]
        public void NormalizedCVectorResultantTest()
        {
            // tests resultant normalization
            var dag = DAGWithMultipleFormulasAndConstants();
            var wbname = dag.getWorkbookName();
            var wsname = dag.getWorksheetNames()[0];
            var path = dag.getWorkbookDirectory();
            var formula_b1 = AST.Address.fromA1withMode(1, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var formula_b2 = AST.Address.fromA1withMode(2, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var formula_b3 = AST.Address.fromA1withMode(3, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);
            var formula_b4 = AST.Address.fromA1withMode(4, "B", AST.AddressMode.Relative, AST.AddressMode.Relative, wsname, wbname, path);

            var resultant_b1 = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula_b1, dag);
            var resultant_b1_shouldbe = Resultant.NewCVectorResultant(-1, 0, 0, 3);
            Assert.AreEqual(resultant_b1_shouldbe, resultant_b1);

            var resultant_b2 = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula_b2, dag);
            var resultant_b2_shouldbe = Resultant.NewCVectorResultant(-1, 0, 0, 1);
            Assert.AreEqual(resultant_b2_shouldbe, resultant_b2);

            var resultant_b3 = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula_b3, dag);
            var resultant_b3_shouldbe = Resultant.NewCVectorResultant(-1, 0, 0, 0);
            Assert.AreEqual(resultant_b3_shouldbe, resultant_b3);

            var resultant_b4 = (Resultant)ExceLint.Vector.ShallowInputVectorMixedCVectorResultantNotOSI.run(formula_b4, dag);
            var resultant_b4_shouldbe = Resultant.NewCVectorResultant(-2, -3, 0, 8);
            Assert.AreEqual(resultant_b4_shouldbe, resultant_b4);

            Resultant[] rs = { resultant_b1, resultant_b2, resultant_b3, resultant_b4 };
            ExceLint.Countable[] rs_normalized = ExceLint.Countable.Normalize(rs);
            ExceLint.Countable rs_normalized_b1 = Resultant.NewCVectorResultant(1.0, 1.0, 0.0, 0.375);
            ExceLint.Countable rs_normalized_b2 = Resultant.NewCVectorResultant(1.0, 1.0, 0.0, 0.125);
            ExceLint.Countable rs_normalized_b3 = Resultant.NewCVectorResultant(1.0, 1.0, 0.0, 0.0);
            ExceLint.Countable rs_normalized_b4 = Resultant.NewCVectorResultant(0.0, 0.0, 0.0, 1.0);
            ExceLint.Countable[] rs_normalized_shouldbe = { rs_normalized_b1, rs_normalized_b2, rs_normalized_b3, rs_normalized_b4 };

            Assert.AreEqual(rs_normalized[0], rs_normalized_b1);
            Assert.AreEqual(rs_normalized[1], rs_normalized_b2);
            Assert.AreEqual(rs_normalized[2], rs_normalized_b3);
            Assert.AreEqual(rs_normalized[3], rs_normalized_b4);
        }
    }
}
