using System;
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
                                new CRTLeaf<string>(UInt128.Zero.Sub(UInt128.One), "all one or none!")
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
                                new CRTLeaf<string>(UInt128.Zero.Sub(UInt128.One), "all one or none!")
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
            var key = UInt128.Zero.Sub(UInt128.One);
            var value = t.Lookup(key);

            Assert.IsTrue(FSharpOption<string>.get_IsSome(value));
            Assert.AreEqual("all one or none!", value.Value);
        }

        [TestMethod]
        public void SubtreeLookupTest()
        {
            // initialize tree
            var t = setupTree();

            // lookup a subtree; should be the entire tree
            var zero = UInt128.Zero;
            var st_opt = t.LookupSubtree(zero, zero);

            // the query should have returned a tree
            Assert.IsTrue(FSharpOption<CRTNode<string>>.get_IsSome(st_opt));

            var st = st_opt.Value;

            // the returned tree should be the root
            Assert.AreEqual(t, st);
        }

        [TestMethod]
        public void SubtreeLookupTest2()
        {
            // initialize tree
            var t = setupTree();

            // lookup a subtree; should be the entire tree
            var key = UInt128.Zero;
            var mask = UInt128.calcMask(0, 1);
            var st_opt = t.LookupSubtree(key, mask);

            // the query should have returned a tree
            Assert.IsTrue(FSharpOption<CRTNode<string>>.get_IsSome(st_opt));

            var st = st_opt.Value;

            // the returned tree should be an inner node with no value
            Assert.AreEqual(FSharpOption<string>.None, st.Value);

            // the returned tree should have a left subtree
            Assert.IsTrue(st.GetType() == typeof(CRTInner<string>));

            var st_l = ((CRTInner<string>)st).Left;

            // the left subtree should have a left subtree
            Assert.IsTrue(st_l.GetType() == typeof(CRTInner<string>));

            var st_ll = ((CRTInner<string>)st_l).Left;

            // the left left subtree should be a leaf
            Assert.IsTrue(st_ll.GetType() == typeof(CRTLeaf<string>));

            // the left left subtree value should be "hi!"
            var st_ll_value = ((CRTLeaf<string>)st_ll).Value;
            Assert.IsTrue(FSharpOption<string>.get_IsSome(st_ll_value));
            Assert.AreEqual("hi!", st_ll_value.Value);
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
