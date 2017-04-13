namespace ExceLint

    open System.Collections.Generic
    open UInt128

    module CRTUtil =
        let nBitMask(n: int) : UInt128 =
            UInt128.Sub (UInt128.LeftShift UInt128.One n) UInt128.One

        let calcMask(startpos: int)(endpos: int) : UInt128 =
            UInt128.LeftShift (nBitMask(endpos - startpos + 1)) (127 - startpos)

    [<AbstractClass>]
    // endpos is inclusive
    type CRTNode<'a>(endpos: int, prefix: UInt128) =
        abstract member IsLeaf: bool
        abstract member IsEmpty: bool
        abstract member Lookup: UInt128 -> 'a option
        
    and CRTRoot<'a>(left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(-1, UInt128.Zero)
        let topbit = UInt128.LeftShift UInt128.One 127
        override self.IsLeaf = false
        override self.IsEmpty = false
        override self.Lookup(key: UInt128) : 'a option =
            // is the higest-order bit 0 or 1?
            if UInt128.GreaterThan topbit key then
                // top bit is 0
                left.Lookup key
            else
                // top bit is 1
                right.Lookup key

    and CRTInner<'a>(endpos: int, prefix: UInt128, left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(endpos, prefix)
        let mask = CRTUtil.calcMask 0 endpos
        let mybits = UInt128.BitwiseAnd mask prefix
        let nextBitMask = CRTUtil.calcMask (endpos + 1) (endpos + 1)
        override self.IsLeaf = false
        override self.IsEmpty = false
        override self.Lookup(key: UInt128) : 'a option =
            let keybits = UInt128.BitwiseAnd mask key
            if UInt128.Equals mybits keybits then
                let nextbit = UInt128.BitwiseAnd nextBitMask key
                if UInt128.GreaterThan nextBitMask nextbit then
                    left.Lookup key
                else
                    right.Lookup key
            else
                None

    and CRTLeaf<'a>(prefix: UInt128, value: 'a) =
        inherit CRTNode<'a>(127, prefix)
        override self.IsLeaf = true
        override self.IsEmpty = false
        override self.Lookup(str: UInt128) : 'a option = Some value

    and CRTEmptyLeaf<'a>(prefix: UInt128) =
        inherit CRTNode<'a>(127, prefix)
        override self.IsLeaf = true
        override self.IsEmpty = true
        override self.Lookup(str: UInt128) : 'a option = None