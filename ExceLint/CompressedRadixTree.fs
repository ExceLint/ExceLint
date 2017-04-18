namespace ExceLint

    open System.Collections.Generic

    [<AbstractClass>]
    // endpos is inclusive
    type CRTNode<'a when 'a : equality>(endpos: int, prefix: UInt128) =
        abstract member IsLeaf: bool
        abstract member IsEmpty: bool
        abstract member Lookup: UInt128 -> 'a option
        abstract member LookupSubtree: UInt128 -> UInt128 -> CRTNode<'a> option
        abstract member Replace: UInt128 -> 'a -> CRTNode<'a>
        abstract member Value: 'a option
        abstract member SubtreeValues: UInt128 -> UInt128 -> seq<'a>
        abstract member LRTraversal: seq<'a>
        
    and CRTRoot<'a when 'a : equality>(left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(-1, UInt128.Zero)
        let topbit = UInt128.One.LeftShift 127
        new() = CRTRoot(CRTEmptyLeaf(UInt128.Zero),CRTEmptyLeaf(UInt128.Zero.Sub(UInt128.One)))
        member self.Left = left
        member self.Right = right
        override self.IsLeaf = false
        override self.IsEmpty = false
        override self.Value = None
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option =
            if mask = UInt128.Zero then
                // return entire tree
                Some (self :> CRTNode<'a>)
            else
                // return subtree
                if topbit.GreaterThan key then
                    left.LookupSubtree key mask
                else
                    right.LookupSubtree key mask
        override self.Lookup(key: UInt128) : 'a option =
            match self.LookupSubtree key (UInt128.Zero.Sub(UInt128.One)) with
            | Some st -> st.Value
            | None -> None
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            if topbit.GreaterThan key then
                // top bit is 0, replace left
                CRTRoot(left.Replace key value, right) :> CRTNode<'a>
            else
                // top bit is 1, replace right
                CRTRoot(left, right.Replace key value) :> CRTNode<'a>
        override self.SubtreeValues(key: UInt128)(value: UInt128) : seq<'a> =
            match self.LookupSubtree key value with
            | Some(st) -> st.LRTraversal
            | None -> Seq.empty
        override self.LRTraversal: seq<'a> =
            // root nodes store no data
            Seq.concat [ self.Left.LRTraversal; self.Right.LRTraversal ]
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTRoot<'a> as other ->
                self.Left = other.Left &&
                self.Right = other.Right
            | _ -> false
        override self.GetHashCode() : int =
            self.Left.GetHashCode() ^^^ self.Right.GetHashCode()

    and CRTInner<'a when 'a : equality>(endpos: int, prefix: UInt128, left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(endpos, prefix)
        let mymask = UInt128.calcMask 0 endpos
        let mybits = mymask.BitwiseAnd prefix
        let nextBitMask = UInt128.calcMask (endpos + 1) (endpos + 1)
        member self.Left = left
        member self.Right = right
        member self.PrefixLength = endpos + 1
        member self.Prefix = prefix
        override self.IsLeaf = false
        override self.IsEmpty = false
        override self.Value = None
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option = 
            // the user wants this tree, or
            // the user wants a nonexistent parent tree
            if mask = mymask || (mask.BitwiseAnd mymask) = mask then
                Some (self :> CRTNode<'a>)
            else
            // the user wants a subtree
                let keybits = mymask.BitwiseAnd key
                if mybits = keybits then
                    let nextbit = nextBitMask.BitwiseAnd key
                    if nextBitMask.GreaterThan nextbit then
                        left.LookupSubtree key mask
                    else
                        right.LookupSubtree key mask
                else
                // the requested subtree is not in the tree
                    None
        override self.Lookup(key: UInt128) : 'a option =
            match self.LookupSubtree key (UInt128.Zero.Sub(UInt128.One)) with
            | Some st -> st.Value
            | None -> None
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            let keybits = mymask.BitwiseAnd key
            if mybits = keybits then
                let nextbit = nextBitMask.BitwiseAnd key
                if nextBitMask.GreaterThan nextbit then
                    CRTInner(endpos, prefix, left.Replace key value, right) :> CRTNode<'a>
                else
                    CRTInner(endpos, prefix, left, right.Replace key value) :> CRTNode<'a>
            else
                // insert a new parent
                // find longest common prefix
                let pidx = key.LongestCommonPrefix prefix
                let mask' = UInt128.calcMask 0 (endpos - 1)
                let prefix' = mask'.BitwiseAnd key

                // insert current subtree on the left or on the right of new parent node?
                let nextBitMask' = UInt128.calcMask pidx pidx
                let nextbit = nextBitMask'.BitwiseAnd prefix
                if nextBitMask'.GreaterThan nextbit then
                    // current node goes on the left
                    CRTInner(pidx - 1, prefix', self, CRTLeaf(key, value)) :> CRTNode<'a>
                else
                    // current node goes on the right
                    CRTInner(pidx - 1, prefix', CRTLeaf(key, value), self) :> CRTNode<'a>
        override self.SubtreeValues(key: UInt128)(value: UInt128) : seq<'a> =
            match self.LookupSubtree key value with
            | Some(st) -> st.LRTraversal
            | None -> Seq.empty
        override self.LRTraversal: seq<'a> =
            // inner nodes store no data
            Seq.concat [ self.Left.LRTraversal; self.Right.LRTraversal ]
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTInner<'a> as other ->
                self.Left = other.Left &&
                self.Right = other.Right &&
                self.PrefixLength = other.PrefixLength &&
                self.Prefix = other.Prefix
            | _ -> false
        override self.GetHashCode() : int =
            self.Left.GetHashCode() ^^^ self.Right.GetHashCode()

    and CRTLeaf<'a when 'a : equality>(prefix: UInt128, value: 'a) =
        inherit CRTNode<'a>(127, prefix)
        member self.Prefix = prefix
        override self.IsLeaf = true
        override self.IsEmpty = false
        override self.Value = Some value
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option = 
            Some (self :> CRTNode<'a>)
        override self.Lookup(str: UInt128) : 'a option = Some value
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            CRTLeaf(prefix, value) :> CRTNode<'a>
        override self.SubtreeValues(key: UInt128)(value: UInt128) : seq<'a> = self.LRTraversal
        override self.LRTraversal: seq<'a> = seq [ value ]
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTLeaf<'a> as other ->
                self.Prefix = other.Prefix &&
                self.Value = other.Value
            | _ -> false
        override self.GetHashCode() : int =
            prefix.GetHashCode()

    and CRTEmptyLeaf<'a when 'a : equality>(prefix: UInt128) =
        inherit CRTNode<'a>(127, prefix)
        member self.Prefix = prefix
        override self.IsLeaf = true
        override self.IsEmpty = true
        override self.Value = None
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option = None
        override self.Lookup(str: UInt128) : 'a option = None
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            CRTLeaf(prefix, value) :> CRTNode<'a>
        override self.SubtreeValues(key: UInt128)(value: UInt128) : seq<'a> = self.LRTraversal
        override self.LRTraversal: seq<'a> = Seq.empty
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTEmptyLeaf<'a> as other ->
                self.Prefix = other.Prefix
            | _ -> false
        override self.GetHashCode() : int =
            prefix.GetHashCode()