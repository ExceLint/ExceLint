using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SquareVector = ExceLint.Vector.SquareVector;
using Edge = System.Tuple<ExceLint.Vector.SquareVector, ExceLint.Vector.SquareVector>;
using DistDict = System.Collections.Generic.Dictionary<System.Tuple<ExceLint.Vector.SquareVector, ExceLint.Vector.SquareVector>, double>;
using System.Collections.Generic;
using ExceLint;

namespace ExceLintTests
{
    [TestClass]
    public class COFTests
    {
        static SquareVector p1 = new SquareVector(0, 0, 11.56, 8.75);

        static SquareVector nn1 = new SquareVector(0, 0, 1, 1);     // no name
        static SquareVector nn2 = new SquareVector(0, 0, 2, 1);     // no name
        static SquareVector nn3 = new SquareVector(0, 0, 3, 1);     // no name
        static SquareVector nn4 = new SquareVector(0, 0, 4, 1);     // no name
        static SquareVector nn5 = new SquareVector(0, 0, 5, 1);     // no name
        static SquareVector o3 = new SquareVector(0, 0, 6, 1);      // no name
        static SquareVector o4 = new SquareVector(0, 0, 7, 1);      // no name
        static SquareVector o5 = new SquareVector(0, 0, 8, 1);      // no name
        static SquareVector o6 = new SquareVector(0, 0, 9, 1);      // 6
        static SquareVector o7 = new SquareVector(0, 0, 10, 1);     // 7
        static SquareVector o8 = new SquareVector(0,0,11,1);        // 8
        static SquareVector o9 = new SquareVector(0,0,12,1);        // 9
        static SquareVector o10 = new SquareVector(0,0,13,1);       // 10
        static SquareVector o11 = new SquareVector(0,0,14,1);       // 11
        static SquareVector o12 = new SquareVector(0,0,15,1);       // 12
        static SquareVector o13 = new SquareVector(0,0,16,1);       // 13
        static SquareVector o14 = new SquareVector(0,0,17,1);       // 14
        static SquareVector nn18 = new SquareVector(0,0,18,1);      // no name
        static SquareVector nn19 = new SquareVector(0,0,19,1);      // no name
        static SquareVector nn20 = new SquareVector(0,0,20,1);      // no name
        static SquareVector nn21 = new SquareVector(0,0,21,1);      // no name
        static SquareVector o2 = new SquareVector(0,0,10,4);        // 2
        static SquareVector o1 = p1;                                // 1

        static SquareVector[] input_arr =
            { nn1, nn2, nn3, nn4, nn5, o3, o4, o5, o6, o7, o8, o9,
              o10, o11, o12, o13, o14, nn18, nn19, nn20, nn21, o2, o1 };

        static SquareVector[] k10_arr =
            { o2, o9, o10, o8, o11, o7, o12, o6, o13, o5 };

        HashSet<SquareVector> input = new HashSet<SquareVector>(input_arr);

        HashSet<SquareVector> k10 = new HashSet<SquareVector>(k10_arr);

        DistDict dd = new DistDict(Vector.pairwiseDistances(Vector.edges(input_arr)));

        private HashSet<T> SymmetricDiff<T>(HashSet<T> hs1, HashSet<T> hs2)
        {
            var retVal = new HashSet<T>(hs1);
            retVal.SymmetricExceptWith(hs2);
            return retVal;
        }

        [TestMethod]
        public void COFPaperExamplePoint1KNeighbors()
        {
            var kN = Vector.Nk(p1, 10, input, dd);
            HashSet<SquareVector> sd = SymmetricDiff(k10, kN);
            Assert.IsTrue(sd.Count == 0);
        }

        [TestMethod]
        public void COFPaperExamplePoint1SBNTrail()
        {
            // ensures that the SBN trail produced by ExceLint's COF algorithm
            // is faithful to the output shown in the paper

            // expected output
            Edge[] expected_path = {
                new Edge(o1, o2),
                new Edge(o2, o7),
                new Edge(o7, o6),
                new Edge(o6, o5),
                new Edge(o7, o8),
                new Edge(o8, o9),
                new Edge(o9, o10),
                new Edge(o10,o11),
                new Edge(o11,o12),
                new Edge(o12,o13)
            };

            // get k neighbors
            var kN = Vector.Nk(p1, 10, input, dd);

            // actual output
            Edge[] path = Vector.SBNTrail(p1, kN, dd);

            Assert.IsTrue(Enumerable.SequenceEqual(path, expected_path));
        }
    }
}
