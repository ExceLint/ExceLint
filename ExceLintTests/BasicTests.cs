using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using COMWrapper;

namespace ExceLintTests
{
    [TestClass]
    public class BasicTests
    {
        Depends.DAG _addressModeDAG;

        public BasicTests()
        {
            _addressModeDAG = AddressModeDAG();
        }

        [TestMethod]
        [Ignore]    // painfully slow
        public void canBuildDependenceGraph()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\7_gradebook_xls.xlsx");
            var graph = wb.buildDependenceGraph();
            var dotty = graph.ToDOT();
            Assert.IsTrue(dotty.Length != 0);
        }

        private Depends.DAG AddressModeDAG()
        {
            var app = new Application();
            var wb = app.OpenWorkbook(@"..\..\TestData\AddressModes.xlsx");
            var graph = wb.buildDependenceGraph();
            return graph;
        }

        [TestMethod]
        public void absoluteSingleVector()
        {
            // tests that $A$2 in cell B2 on the same sheet returns the vector (1,2,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var cell = AST.Address.fromA1withMode(2, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = ExceLint.Vector.getVectors(cell, _addressModeDAG, transitive: false, isForm: true, isRel: true, isMixed: true);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(1, 2, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void relativeRowAbsoluteColumnSingleVector()
        {
            // tests that A$3 in cell B3 on the same sheet returns the vector (-1,3,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var cell = AST.Address.fromA1withMode(3, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = ExceLint.Vector.getVectors(cell, _addressModeDAG, transitive: false, isForm: true, isRel: true, isMixed: true);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(-1, 3, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void absoluteRowRelativeColumnSingleVector()
        {
            // tests that $A4 in cell B4 on the same sheet returns the vector (1,0,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var cell = AST.Address.fromA1withMode(4, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = ExceLint.Vector.getVectors(cell, _addressModeDAG, transitive: false, isForm: true, isRel: true, isMixed: true);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(1, 0, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void relativeSingleVector()
        {
            // tests that A5 in cell B5 on the same sheet returns the vector (-1,0,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var cell = AST.Address.fromA1withMode(5, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = ExceLint.Vector.getVectors(cell, _addressModeDAG, transitive: false, isForm: true, isRel: true, isMixed: true);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(-1, 0, 0);
            Assert.IsTrue(foo.Equals(bar));
        }
    }
}
