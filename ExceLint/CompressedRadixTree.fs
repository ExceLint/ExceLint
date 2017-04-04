namespace ExceLint

    open System.Collections.Generic
    open UInt128

    module CRTUtil =
        let nBitMask(n: int) : UInt128 =
            UInt128.Sub (UInt128.LeftShift UInt128.One n) UInt128.One

        let calcMask(startpos: int)(endpos: int) : UInt128 =
            UInt128.LeftShift (nBitMask(endpos - startpos + 1)) (127 - startpos)

    [<AbstractClass>]
    // start and end are inclusive
    type CRTNode(startpos: int, endpos: int, prefix: UInt128) =
        abstract member isLeaf: bool
        abstract member isEmpty: bool
        abstract member lookup: UInt128 -> CRTNode
        
    and CRTInner(startpos: int, endpos: int, prefix: UInt128, left: CRTNode, right: CRTNode) =
        inherit CRTNode(startpos, endpos, prefix)
        let mask = CRTUtil.calcMask startpos endpos
        let mybits = UInt128.BitwiseAnd mask prefix
        let nextBitMask = CRTUtil.calcMask (endpos + 1) (endpos + 1)
        override self.isLeaf = false
        override self.isEmpty = false
        override self.lookup(str: UInt128) : CRTNode =
            let strbits = UInt128.BitwiseAnd mask str
            if UInt128.Equals mybits strbits then
                // TODO still need to right shift the following:
                let nextbit = UInt128.BitwiseAnd nextBitMask str
                // if 0 return left child
                // if 1 return right child
                failwith "not yet"
            else
                CRTEmptyLeaf(str) :> CRTNode

    and CRTLeaf(startpos: int, endpos: int, prefix: UInt128, points: HashSet<AST.Address>) =
        inherit CRTNode(startpos, endpos, prefix)
        override self.isLeaf = true
        override self.isEmpty = false
        override self.lookup(str: UInt128) : CRTNode = failwith "not yet"

    and CRTEmptyLeaf(prefix: UInt128) =
        inherit CRTNode(0, 127, prefix)
        override self.isLeaf = true
        override self.isEmpty = true
        override self.lookup(str: UInt128) : CRTNode = self :> CRTNode