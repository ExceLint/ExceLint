using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExceLintRunnerTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void CUSTODESOutputParseTest1()
        {
            var output = "----procesing worksheet 'Figure II.10'----";
            var r = CUSTODESGrammar.parse(output);

            if (r.IsCSuccess)
            {
                var results = ((CUSTODESGrammar.CUSTODESParse.CSuccess)r).Item;
                Assert.IsTrue(results.ContainsKey("Figure II.10"));
                Assert.IsTrue(results["Figure II.10"].Length == 0);
            } else
            {
                var error = ((CUSTODESGrammar.CUSTODESParse.CFailure)r).Item;
                Assert.Fail(error);
            }
        }

        [TestMethod]
        public void CUSTODESOutputParseTest2()
        {
            var output = Properties.Resources.CUSTODESOutput;
            var r = CUSTODESGrammar.parse(output);

            if (r.IsCSuccess)
            {
                var results = ((CUSTODESGrammar.CUSTODESParse.CSuccess)r).Item;
                Assert.IsTrue(results.Count == 24);
                Assert.IsTrue(results.ContainsKey("Table II.5") && results["Table II.5"].Length == 19);
                Assert.IsTrue(results.ContainsKey("Figure II.10") && results["Figure II.10"].Length == 0);
            }
            else
            {
                var error = ((CUSTODESGrammar.CUSTODESParse.CFailure)r).Item;
                Assert.Fail(error);
            }
        }

        [TestMethod]
        public void CUSTODESExceptionTest()
        {
            var output = "----procesing worksheet 'Budget Process'----" +
                         "----procesing worksheet 'Transf Plan'----" +
                         "----procesing worksheet 'Policies'----" +
                         "----procesing worksheet 'Income Chart'----" +
                         "----procesing worksheet 'Expense Chart'----" +
                         "----procesing worksheet 'General PTSA  Budget '----" +
                         "---- Stage I clustering begined ----" +
                         "----Stage I finished ----" +
                         "found 18 clusters" +
                         "---- Stage II begined----" +
                         "detected 6 smells:" +
                         "            I113" +
                         "            H113" +
                         "G9" +
                         "J22" +
                         "I66" +
                         "I76" +
                         "---Analysis Finished-- -" +
                         "----procesing worksheet 'EDK budget'----" +
                         "---- Stage I clustering begined ----" +
                         "----Stage I finished ----" +
                         "found 8 clusters" +
                         "---- Stage II begined----" +
                         "detected 0 smells:" +
                         "            ---Analysis Finished-- -" +
                         "            ----procesing worksheet 'collapsed budget for charts'----" +
                         "            ---- Stage I clustering begined ----" +
                         "            Exception in thread \"main\" java.lang.OutOfMemoryError: Java heap space" +
                         "\tat distance.RTED_InfoTree_Opt.init(RTED_InfoTree_Opt.java:154)" +
                         "\tat distance.RTED_InfoTree_Opt.nonNormalizedTreeDist(RTED_InfoTree_Opt.java:123)" +
                         "\tat convenience.RTED.computeDistance(RTED.java:42)" +
                         "\tat b.b.a(Unknown Source)" +
                         "\tat a.b.g.a(Unknown Source)" +
                         "\tat spreadsheet.cc2.Main.main(Unknown Source)";
            var r = CUSTODESGrammar.parse(output);

            if (r.IsCSuccess)
            {
                Assert.Fail("This test should never succeed.");
                
            }
            else
            {
                var error = ((CUSTODESGrammar.CUSTODESParse.CFailure)r).Item;
                Assert.IsTrue(error.StartsWith("Exception in thread"));
            }
        }
    }
}
