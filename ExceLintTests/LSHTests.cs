using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Numerics;

namespace ExceLintTests
{
    [TestClass]
    public class LSHTests
    {
        [TestMethod]
        public void SimpleHashTest()
        {
            var x = 1;
            var y = 1;
            var z = 1;
            var bv = Convert.ToUInt64(0x10000000003);
            var bv_is = ExceLint.LSHCalc.hashi(x, y, z);
            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void LessSimpleHashTest()
        {
            var x = 37;                      // 100101
            var y = 21;                      // 010101
            var z = 16777215;                // 111111111111111111111111
            var bv = 18446742974197925427UL; // 111111111111111111111111000000000000000000000000000011000110011
            var bv_is = ExceLint.LSHCalc.hashi(x, y, z);
            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void SimpleCountableHashTest()
        {
            var o = BigInteger.One;
            var c = ExceLint.FeatureUtil.makeFullCVR(1, 1, 1, 1, 1, 1, 1);
            var bv = o | (o << 1) | (o << 2) | (o << 3) | (o << 4) | (o << (20 * 5)) | (o << (20 * 5 + 1));
            var bv_is = ExceLint.LSHCalc.h7(c);

            // helpful for debugging...
            //var pp_bv = ExceLint.LSHCalc.ppbi(bv);
            //var pp_bv_is = ExceLint.LSHCalc.ppbi(bv_is);

            Assert.AreEqual(bv, bv_is);
        }
    }
}
