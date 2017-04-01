using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Numerics;
using ExceLint;

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
            var bv_is = LSHCalc.hashi(x, y, z);
            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void LessSimpleHashTest()
        {
            var x = 37;                      // 100101
            var y = 21;                      // 010101
            var z = 16777215;                // 111111111111111111111111
            var bv = 18446742974197925427UL; // 111111111111111111111111000000000000000000000000000011000110011
            var bv_is = LSHCalc.hashi(x, y, z);
            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void SimpleCountableHashTest()
        {
            var o = BigInteger.One;
            var c = FeatureUtil.makeFullCVR(1, 1, 1, 1, 1, 1, 1);
            var bvbi = o | (o << 1) | (o << 2) | (o << 3) | (o << 4) | (o << (20 * 5)) | (o << (20 * 5 + 1));
            var bv = UInt128.FromBigInteger(bvbi);
            var bv_is = LSHCalc.h7(c);

            var bv_pp = UInt128.prettyPrint(bv);
            var bv_is_pp = UInt128.prettyPrint(bv_is);

            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void Uint128LeftShiftLowOverflowTest()
        {
            var n = new UInt128.UInt128(0UL, UInt64.MaxValue);
            var n2 = UInt128.LeftShift(n, 1);

            var ppn = UInt128.prettyPrint(n);
            var ppn2 = UInt128.prettyPrint(n2);
            var ppnsb = UInt128.prettyPrint(new UInt128.UInt128(0UL, 9223372036854775806UL));

            Assert.AreEqual(1UL, n2.High);
            Assert.AreEqual(18446744073709551614UL, n2.Low);
        }

        [TestMethod]
        public void Uint128LeftShiftLowOverflowTest2()
        {
            var n = new UInt128.UInt128(0UL, 1UL);
            var n2 = UInt128.LeftShift(n, 100);
            var sb = UInt128.FromBigInteger(BigInteger.One << 100);

            var ppn = UInt128.prettyPrint(n);
            var ppn2 = UInt128.prettyPrint(n2);
            var ppnsb = UInt128.prettyPrint(sb);

            Assert.AreEqual(sb, n2);
        }

        [TestMethod]
        public void UInt128ToBigDecimalAndBackTest()
        {
            var bignum = new UInt128.UInt128(UInt64.MaxValue, UInt64.MaxValue);
            var bignumbi = UInt128.ToBigInteger(bignum);
            var bignum2 = UInt128.FromBigInteger(bignumbi);

            Assert.AreEqual(bignum, bignum2);
        }

        [TestMethod]
        public void BigIntegerToUInt128AndBackTest()
        {
            var bignum = new BigInteger(UInt64.MaxValue * Convert.ToUInt64(20));
            var bignum128 = UInt128.FromBigInteger(bignum);
            var bignum2 = UInt128.ToBigInteger(bignum128);

            Assert.AreEqual(bignum, bignum2);
        }

        [TestMethod]
        public void CountOnesTest()
        {
            var n = new UInt128.UInt128(1UL, 1UL);
            var i = UInt128.CountOnes(n);
            Assert.AreEqual(2, i);
        }

        [TestMethod]
        public void CountZeroesTest()
        {
            var n = new UInt128.UInt128(3UL, 3UL);
            var i = UInt128.CountZeroes(n);
            Assert.AreEqual(124, i);
        }

        [TestMethod]
        public void CommonPrefixTest()
        {
            var n = new UInt128.UInt128(123456789UL, 123456789UL);
            var n2 = new UInt128.UInt128(123456789UL, ~123456789UL);
            var len = UInt128.longestCommonPrefix(n, n2);
            Assert.AreEqual(64, len);
        }
    }
}
