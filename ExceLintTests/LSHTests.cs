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

            Assert.AreEqual(bv, bv_is);
        }

        [TestMethod]
        public void Uint128LeftShiftLowOverflowTest()
        {
            var n = new UInt128(0UL, UInt64.MaxValue);
            var n2 = n.LeftShift(1);

            Assert.AreEqual(1UL, n2.High);
            Assert.AreEqual(18446744073709551614UL, n2.Low);
        }

        [TestMethod]
        public void Uint128LeftShiftLowOverflowTest2()
        {
            var n = new UInt128(0UL, 1UL);
            var n2 = n.LeftShift(100);
            var sb = UInt128.FromBigInteger(BigInteger.One << 100);

            Assert.AreEqual(sb, n2);
        }

        [TestMethod]
        public void UInt128LeftShiftOverflowTest3()
        {
            var n = UInt128.FromBigInteger(BigInteger.Parse("18446744073709551616"));
            var n2 = n.LeftShift(1);
            var sb = UInt128.FromBigInteger(BigInteger.Parse("36893488147419103232"));
            Assert.AreEqual(sb, n2);
        }

        [TestMethod]
        public void UInt128LeftShiftTest()
        {
            for (int i = 0; i < 128; i++)
            {
                var n = UInt128.One.LeftShift(i);
                var sb = UInt128.FromBigInteger(BigInteger.Pow(new BigInteger(2), i));
                if (!sb.Equals(n))
                {
                    Console.WriteLine("Whoa!");
                }
                Assert.AreEqual(sb, n);
            }
        }

        [TestMethod]
        public void UInt128ToBigDecimalAndBackTest()
        {
            var bignum = new UInt128(UInt64.MaxValue, UInt64.MaxValue);
            var bignumbi = bignum.ToBigInteger;
            var bignum2 = UInt128.FromBigInteger(bignumbi);

            Assert.AreEqual(bignum, bignum2);
        }

        [TestMethod]
        public void BigIntegerToUInt128AndBackTest()
        {
            var bignum = new BigInteger(UInt64.MaxValue * Convert.ToUInt64(20));
            var bignum128 = UInt128.FromBigInteger(bignum);
            var bignum2 = bignum128.ToBigInteger;

            Assert.AreEqual(bignum, bignum2);
        }

        [TestMethod]
        public void CountOnesTest()
        {
            var n = new UInt128(1UL, 1UL);
            var i = n.CountOnes;
            Assert.AreEqual(2, i);
        }

        [TestMethod]
        public void CountZeroesTest()
        {
            var n = new UInt128(3UL, 3UL);
            var i = n.CountZeroes;
            Assert.AreEqual(124, i);
        }

        [TestMethod]
        public void CommonPrefixTest()
        {
            var n = new UInt128(UInt64.MaxValue, UInt64.MaxValue);
            var n2 = new UInt128(UInt64.MaxValue, ~ UInt64.MaxValue);
            var len = n.LongestCommonPrefix(n2);
            Assert.AreEqual(64, len);
        }

        [TestMethod]
        public void CommonPrefixTest2()
        {
            var n = new UInt128(UInt64.MaxValue, 0UL);
            var n2 = new UInt128(UInt64.MaxValue, 3UL);
            var len = n.LongestCommonPrefix(n2);
            Assert.AreEqual(126, len);
        }

        [TestMethod]
        public void CommonPrefixTest3()
        {
            var a = UInt128.FromBinaryString("11111010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010");
            var b = UInt128.FromBinaryString("11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101");
            var c = BigInteger.Parse("326103934965899360819067332122111202645");
            var len = a.LongestCommonPrefix(b);
            Assert.AreEqual(4, len);
        }

        [TestMethod]
        public void XORTest()
        {
            var a = new UInt128(0UL, 0UL);
            var b = UInt128.MaxValue;
            var c = a.BitwiseXor(b);
            Assert.AreEqual(UInt128.MaxValue, c);
        }

        [TestMethod]
        public void XORTest2()
        {
            var a = new UInt128(0UL, UInt64.MaxValue);
            var b = new UInt128(UInt64.MaxValue, 0UL);
            var c = a.BitwiseXor(b);
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
            var abi = a.ToBigInteger;

            var bbi = BigInteger.Parse("326103934965899360819067332122111202645");
            var b = UInt128.FromBigInteger(bbi);

            Assert.IsTrue(UInt128.Equals(a, b));
            Assert.AreEqual(bbi, abi);
        }

        [TestMethod]
        public void UInt128ToBinaryStringAndBack()
        {
            var b1 = BigInteger.Parse("123456789123456789");
            var n1 = UInt128.FromBigInteger(b1);
            var s = n1.ToString();
            var n2 = UInt128.FromBinaryString(s);
            var b2 = n2.ToBigInteger;
            Assert.AreEqual(b1, b2);
        }

        [TestMethod]
        public void UInt128AddToZeroTest()
        {
            var result = UInt128.One.Add(UInt128.Zero);
            Assert.AreEqual(UInt128.One, result);
        }

        [TestMethod]
        public void UInt128AddOverflowLow64Test()
        {
            var addend = new UInt128(0UL, UInt64.MaxValue);
            var result = addend.Add(UInt128.One);
            var sb = new UInt128(1UL, 0UL);
            Assert.AreEqual(sb, result);
        }

        [TestMethod]
        public void UInt128SubToZeroTest()
        {
            var result = UInt128.One.Sub(UInt128.One);
            Assert.AreEqual(UInt128.Zero, result);
        }

        [TestMethod]
        public void UInt128SubBorrowHigh64Test()
        {
            var minuend = new UInt128(1UL, 0UL);
            var subtrahend = UInt128.One;
            var result = minuend.Sub(subtrahend);
            var sb = new UInt128(0UL, UInt64.MaxValue);
            Assert.AreEqual(sb, result);
        }

        private UInt128 Reverse(UInt128 n)
        {
            var n_str = n.ToString();
            var n_chars = n_str.ToCharArray();
            Array.Reverse(n_chars);
            var rev_n_str = new String(n_chars);
            return UInt128.FromBinaryString(rev_n_str);
        }

        [TestMethod]
        public void UInt128RightShiftTest()
        {
            var bstr = "11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101";
            var a = UInt128.FromBinaryString(bstr);
            var a_rs1 = a.RightShift(1);

            // get reversed number
            var arev = Reverse(a);
            
            // left shift the reversed number
            var arev_ls1 = arev.LeftShift(1);

            // unreverse shifted number
            var unrev_ls1 = Reverse(arev_ls1);

            // compare
            Assert.AreEqual(a_rs1, unrev_ls1);
        }

        [TestMethod]
        public void UInt128RightShift64Test()
        {
            var bstr = "11110101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101010101";
            var a = UInt128.FromBinaryString(bstr);
            var a_rs64 = a.RightShift(64);

            // get reversed number
            var arev = Reverse(a);

            // left shift the reversed number
            var arev_ls64 = arev.LeftShift(64);

            // unreverse shifted number
            var unrev_ls1 = Reverse(arev_ls64);

            // compare
            Assert.AreEqual(a_rs64, unrev_ls1);
        }
    }
}
