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

        // gets the set of shallow intransitive mixed input vectors pointed to by the formula
        private static Tuple<int,int,int>[] getSIMIVs(AST.Address formula, Depends.DAG dag)
        {
            return ExceLint.Vector.getRebasedVectors(formula, dag, isMixed: true, isTransitive: false, isFormula: true, isOffSheetInsensitive: true, isRelative: true);
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
        public void absoluteSingleVectorSIMIV()
        {
            // tests that $A$2 in cell B2 on the same sheet returns the vector (1,2,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(2, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(1, 2, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void relativeRowAbsoluteColumnSingleVectorSIMIV()
        {
            // tests that A$3 in cell B3 on the same sheet returns the vector (-1,3,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(3, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(-1, 3, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void absoluteRowRelativeColumnSingleVectorSIMIV()
        {
            // tests that $A4 in cell B4 on the same sheet returns the vector (1,0,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(4, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(1, 0, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void relativeSingleVectorSIMIV()
        {
            // tests that A5 in cell B5 on the same sheet returns the vector (-1,0,0)
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(5, "B", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 1);
            var foo = vectors[0];
            var bar = new Tuple<int, int, int>(-1, 0, 0);
            Assert.IsTrue(foo.Equals(bar));
        }

        [TestMethod]
        public void AbsAbsRangeVectorsSIMIV()
        {
            // tests that $A$2:$A$5 in cell C2 on the same sheet returns a set of absolute vectors
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(2, "C", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 4);
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int,int,int>(1, 2, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1, 3, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1, 4, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1, 5, 0))));
        }

        [TestMethod]
        public void RelAbsRangeVectorsSIMIV()
        {
            // tests that A$2:A$5 in cell C3 on the same sheet returns a set of mixed vectors
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(3, "C", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 4);
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, 2, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, 3, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, 4, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, 5, 0))));
        }

        [TestMethod]
        public void AbsRelRangeVectorsSIMIV()
        {
            // tests that $A2:$A5 in cell C4 on the same sheet returns a set of mixed vectors
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(4, "C", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 4);
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1, -2, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1, -1, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1,  0, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(1,  1, 0))));
        }

        [TestMethod]
        public void RelRelRangeVectorsSIMIV()
        {
            // tests that A2:A5 in cell C5 on the same sheet returns a set of mixed vectors
            var wbname = _addressModeDAG.getWorkbookName();
            var wsname = _addressModeDAG.getWorksheetNames()[0];
            var path = _addressModeDAG.getWorkbookDirectory();
            var formula = AST.Address.fromA1withMode(5, "C", AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
            var vectors = getSIMIVs(formula, _addressModeDAG);
            Assert.IsTrue(vectors.Length == 4);
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, -3, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, -2, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2, -1, 0))));
            Assert.IsTrue(Array.Exists(vectors, e => e.Equals(new Tuple<int, int, int>(-2,  0, 0))));
        }
    }
}
