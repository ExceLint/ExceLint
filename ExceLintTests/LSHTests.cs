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
        public void UInt128LeftShiftOverflowTest3()
        {
            var n = UInt128.FromBigInteger(BigInteger.Parse("18446744073709551616"));
            var n2 = UInt128.LeftShift(n, 1);
            var sb = UInt128.FromBigInteger(BigInteger.Parse("36893488147419103232"));
            Assert.AreEqual(sb, n2);
        }

        [TestMethod]
        public void UInt128LeftShiftTest()
        {
            for (int i = 0; i < 128; i++)
            {
                var n = UInt128.LeftShift(UInt128.One, i);
                var sb = UInt128.FromBigInteger(BigInteger.Pow(new BigInteger(2), i));
                if (!UInt128.Equals(sb,n))
                {
                    Console.WriteLine("Whoa!");
                }
                Assert.AreEqual(sb, n);
            }
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
            var n = new UInt128.UInt128(UInt64.MaxValue, UInt64.MaxValue);
            var n2 = new UInt128.UInt128(UInt64.MaxValue, ~ UInt64.MaxValue);
            var len = UInt128.longestCommonPrefix(n, n2);
            Assert.AreEqual(64, len);
        }

        [TestMethod]
        public void CommonPrefixTest2()
        {
            var n = new UInt128.UInt128(UInt64.MaxValue, 0UL);
            var n2 = new UInt128.UInt128(UInt64.MaxValue, 3UL);
            var len = UInt128.longestCommonPrefix(n, n2);
            Assert.AreEqual(126, len);
        }

        [TestMethod]
        public void CommonPrefixTest3()
        {
            var a = UInt128.FromBinaryString("11111010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010");
            var b = UInt128.FromBinaryString("11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101");
            var c = BigInteger.Parse("326103934965899360819067332122111202645");
            var len = UInt128.longestCommonPrefix(a, b);
            Assert.AreEqual(4, len);
        }

        [TestMethod]
        public void XORTest()
        {
            var a = new UInt128.UInt128(0UL, 0UL);
            var b = UInt128.MaxValue;
            var c = UInt128.BitwiseXor(a, b);
            Assert.AreEqual(UInt128.MaxValue, c);
        }

        [TestMethod]
        public void XORTest2()
        {
            var a = new UInt128.UInt128(0UL, UInt64.MaxValue);
            var b = new UInt128.UInt128(UInt64.MaxValue, 0UL);
            var c = UInt128.BitwiseXor(a, b);
            Assert.AreEqual(UInt128.MaxValue, c);
        }

        [TestMethod]
        public void BinaryStringToUInt128Test()
        {
            var s = "10101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010";
            var sb = UInt128.FromBigInteger(BigInteger.Parse("226854911280625642308916404954512140970"));
            var n = UInt128.FromBinaryString(s);
            Assert.AreEqual(sb, n);
        }

        [TestMethod]
        public void BinaryStringToUInt128Test2()
        {
            var a = UInt128.FromBinaryString("11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101");
            var abi = UInt128.ToBigInteger(a);

            var bbi = BigInteger.Parse("326103934965899360819067332122111202645");
            var b = UInt128.FromBigInteger(bbi);

            var ppa = UInt128.prettyPrint(a);
            var ppb = UInt128.prettyPrint(b);

            Assert.IsTrue(UInt128.Equals(a, b));
            Assert.AreEqual(bbi, abi);
        }

        [TestMethod]
        public void UInt128ToBinaryStringAndBack()
        {
            var b1 = BigInteger.Parse("123456789123456789");
            var n1 = UInt128.FromBigInteger(b1);
            var s = UInt128.prettyPrint(n1);
            var n2 = UInt128.FromBinaryString(s);
            var b2 = UInt128.ToBigInteger(n2);
            Assert.AreEqual(b1, b2);
        }

        [TestMethod]
        public void UInt128AddToZeroTest()
        {
            var result = UInt128.Add(UInt128.One, UInt128.Zero);
            Assert.AreEqual(UInt128.One, result);
        }

        [TestMethod]
        public void UInt128AddOverflowLow64Test()
        {
            var addend = new UInt128.UInt128(0UL, UInt64.MaxValue);
            var result = UInt128.Add(addend, UInt128.One);
            var sb = new UInt128.UInt128(1UL, 0UL);
            Assert.AreEqual(sb, result);
        }

        [TestMethod]
        public void UInt128SubToZeroTest()
        {
            var result = UInt128.Sub(UInt128.One, UInt128.One);
            Assert.AreEqual(UInt128.Zero, result);
        }

        [TestMethod]
        public void UInt128SubBorrowHigh64Test()
        {
            var minuend = new UInt128.UInt128(1UL, 0UL);
            var subtrahend = UInt128.One;
            var result = UInt128.Sub(minuend, subtrahend);
            var sb = new UInt128.UInt128(0UL, UInt64.MaxValue);
            Assert.AreEqual(sb, result);
        }

        [TestMethod]
        public void UInt128RightShiftTest()
        {
            var bstr = "11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101";
            var a = UInt128.FromBinaryString(bstr);
            var a_rs1 = UInt128.RightShift(a, 1);

            // get reversed number
            char[] rtsb = bstr.ToCharArray();
            Array.Reverse(rtsb);
            var bstr_rev = new String(rtsb);
            var arev = UInt128.FromBinaryString(bstr_rev);
            
            var arev_ls1 = UInt128.LeftShift(arev, 1);
            Assert.AreEqual(a_rs1, arev_ls1);
        }
    }
}
