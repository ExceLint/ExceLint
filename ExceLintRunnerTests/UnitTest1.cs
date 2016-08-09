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
            var results = CUSTODESGrammar.parse(output);
            Assert.IsTrue(results.ContainsKey("Figure II.10"));
            Assert.IsTrue(results["Figure II.10"].Length == 0);
        }

        [TestMethod]
        public void CUSTODESOutputParseTest2()
        {
            var output = Properties.Resources.CUSTODESOutput;
            var results = CUSTODESGrammar.parse(output);
            Assert.IsTrue(results.Count == 24);
            Assert.IsTrue(results.ContainsKey("Table II.5") && results["Table II.5"].Length == 19);
            Assert.IsTrue(results.ContainsKey("Figure II.10") && results["Figure II.10"].Length == 0);
        }
    }
}
