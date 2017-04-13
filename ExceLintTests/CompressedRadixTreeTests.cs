﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExceLint;
using Microsoft.FSharp.Core;

namespace ExceLintTests
{
    [TestClass]
    public class CompressedRadixTreeTests
    {
        private CRTNode<string> setupTree()
        {
            var t = new CRTRoot<string>(
                        new CRTInner<string>(0, UInt128.FromZeroFilledPrefix("0"),
                                new CRTInner<string>(1, UInt128.FromZeroFilledPrefix("00"),
                                    new CRTInner<string>(4, UInt128.FromZeroFilledPrefix("00001"),
                                        new CRTLeaf<string>(UInt128.FromZeroFilledPrefix("000010"), "hi!"),
                                        new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("000011"))
                                    ),
                                    new CRTInner<string>(3, UInt128.FromZeroFilledPrefix("0010"),
                                        new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("00100")),
                                        new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("00101"))
                                    )
                                ),
                                new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("01"))
                            ),
                            new CRTInner<string>(1, UInt128.FromZeroFilledPrefix("11"),
                                new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("110")),
                                new CRTLeaf<string>(UInt128.Sub(UInt128.Zero, UInt128.One), "all one or none!")
                            )
                        );

            return t;
        }

        private CRTNode<string> treeWithLeftMostInsert()
        {
            var t = new CRTRoot<string>(
                        new CRTInner<string>(0, UInt128.FromZeroFilledPrefix("0"),
                                new CRTInner<string>(1, UInt128.FromZeroFilledPrefix("00"),
                                    new CRTInner<string>(3, UInt128.FromZeroFilledPrefix("0000"),
                                        // START: difference
                                        new CRTLeaf<string>(UInt128.FromZeroFilledPrefix("0000"), "none"),
                                        new CRTInner<string>(4, UInt128.FromZeroFilledPrefix("00001"),
                                            new CRTLeaf<string>(UInt128.FromZeroFilledPrefix("000010"), "hi!"),
                                            new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("000011"))
                                        )
                                        // END: difference
                                    ),
                                    new CRTInner<string>(3, UInt128.FromZeroFilledPrefix("0010"),
                                        new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("00100")),
                                        new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("00101"))
                                    )
                                ),
                                new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("01"))
                            ),
                            new CRTInner<string>(1, UInt128.FromZeroFilledPrefix("11"),
                                new CRTEmptyLeaf<string>(UInt128.FromZeroFilledPrefix("110")),
                                new CRTLeaf<string>(UInt128.Sub(UInt128.Zero, UInt128.One), "all one or none!")
                            )
                        );

            return t;
        }

        [TestMethod]
        public void LookupTest()
        {
            // initialize tree
            var t = setupTree();

            // lookup a value; should be "hi!"
            var key = UInt128.FromZeroFilledPrefix("000010");
            var value = t.Lookup(key);

            Assert.IsTrue(FSharpOption<string>.get_IsSome(value));
            Assert.AreEqual("hi!", value.Value);
        }

        [TestMethod]
        public void LookupTest2()
        {
            // initialize tree
            var t = setupTree();

            // lookup a value; should be "all one or none!"
            var key = UInt128.Sub(UInt128.Zero, UInt128.One);
            var value = t.Lookup(key);

            Assert.IsTrue(FSharpOption<string>.get_IsSome(value));
            Assert.AreEqual("all one or none!", value.Value);
        }

        [TestMethod]
        public void InsertTest()
        {
            // initialize tree
            var t = setupTree();

            // insert a value
            var t2 = t.Replace(UInt128.Zero, "none");

            // expected outcome
            var te = treeWithLeftMostInsert();

            Assert.AreEqual(te, t2);
        }
    }
}
