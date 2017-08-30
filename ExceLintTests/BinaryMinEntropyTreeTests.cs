using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExceLintTests
{
    [TestClass]
    public class BinaryMinEntropyTreeTests
    {
        public AST.Env TestEnvironment()
        {
            var path = "foo";
            var workbook = "bar";
            var worksheet = "foobar";
            return new AST.Env(path, workbook, worksheet);
        }

        public AST.Address AddrWithTestEnv(int row, int col)
        {
            var env = TestEnvironment();
            return AST.Address.fromR1C1withMode(
                row,
                col,
                AST.AddressMode.Absolute,
                AST.AddressMode.Absolute,
                env.WorksheetName,
                env.WorkbookName,
                env.Path);
        }

        [TestMethod]
        public void IsRectDoesNotAcceptNonRectClusters()
        {
            var A3 = AddrWithTestEnv(3, 1);
            var B2 = AddrWithTestEnv(2, 2);
            var B3 = AddrWithTestEnv(3, 2);
            AST.Address[] cArr = {A3, B2, B3};
            var cluster = new HashSet<AST.Address>(cArr);
            var isRect = ExceLint.BinaryMinEntropyTree.ClusterIsRectangular(cluster);
            Assert.IsFalse(isRect);
        }

        [TestMethod]
        public void IsRectDoesNotAcceptNonRectClusters2()
        {
            var A23 = AddrWithTestEnv(23, 1);
            var B23 = AddrWithTestEnv(23, 2);
            var C23 = AddrWithTestEnv(23, 3);
            var B24 = AddrWithTestEnv(24, 2);
            AST.Address[] cArr = { A23, B23, C23, B24 };
            var cluster = new HashSet<AST.Address>(cArr);
            var isRect = ExceLint.BinaryMinEntropyTree.ClusterIsRectangular(cluster);
            Assert.IsFalse(isRect);
        }
    }
}
