namespace ExceLint

    open System.Collections.Generic
    open Utils

    [<AbstractClass>]
    type CRTNode<'a when 'a : equality>(endpos: int, prefix: UInt128) =
        abstract member IsRoot: bool
        abstract member IsLeaf: bool
        abstract member IsEmpty: bool
        abstract member Lookup: UInt128 -> 'a option
        abstract member LookupSubtree: UInt128 -> UInt128 -> CRTNode<'a> option
        abstract member Replace: UInt128 -> 'a -> CRTNode<'a>
        abstract member InsertOr: UInt128 -> 'a -> ('a -> 'a -> 'a) -> CRTNode<'a>
        abstract member Delete: UInt128 -> CRTNode<'a>
        abstract member Value: 'a option
        abstract member EnumerateSubtree: UInt128 -> UInt128 -> seq<'a>
        abstract member LRTraversal: seq<'a>
        abstract member Prefix: UInt128
        abstract member ToGraphViz: string
        abstract member ToGraphVizEdges: GZEdge<'a> list

    and GZEdge<'a when 'a : equality> =
    | Left of CRTNode<'a>*CRTNode<'a>
    | Right of CRTNode<'a>*CRTNode<'a>

    /// <summary>
    /// The root node of a compressed radix tree.
    /// </summary>
    /// <param name="left">
    /// The left subtree.
    /// </param>
    /// <param name="right">
    /// The right subtree.
    /// </param>
    and CRTRoot<'a when 'a : equality>(left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(-1, UInt128.Zero)
        let topbit = UInt128.One.LeftShift 127
        new() = CRTRoot(CRTEmptyLeaf(UInt128.Zero),CRTEmptyLeaf(UInt128.Zero.Sub(UInt128.One)))
        member self.Left = left
        member self.Right = right
        override self.Prefix : UInt128 = failwith "Root has no prefix"
        override self.IsRoot = true
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
        override self.Delete(key: UInt128) : CRTNode<'a> =
            if topbit.GreaterThan key then
                CRTRoot(left.Delete key, right) :> CRTNode<'a>
            else
                CRTRoot(left, right.Delete key) :> CRTNode<'a>
        override self.InsertOr(key: UInt128)(value': 'a)(keyexists: 'a -> 'a -> 'a) : CRTNode<'a> =
            if topbit.GreaterThan key then
                // top bit is 0, replace left
                CRTRoot(left.InsertOr key value' keyexists, right) :> CRTNode<'a>
            else
                // top bit is 1, replace right
                CRTRoot(left, right.InsertOr key value' keyexists) :> CRTNode<'a>
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            self.InsertOr key value (fun i _ -> i)
        override self.EnumerateSubtree(key: UInt128)(value: UInt128) : seq<'a> =
            match self.LookupSubtree key value with
            | Some(st) -> st.LRTraversal
            | None -> Seq.empty
        override self.LRTraversal: seq<'a> =
            // root nodes store no data
            Seq.concat [ self.Left.LRTraversal; self.Right.LRTraversal ]
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTRoot<'a> as other ->
                let ok = self.Left = other.Left &&
                         self.Right = other.Right
                ok
            | _ -> false
        override self.GetHashCode() : int =
            self.Left.GetHashCode() ^^^ self.Right.GetHashCode()
        override self.ToGraphVizEdges : GZEdge<'a> list =
            [
                Left(self, self.Left);
                Right(self, self.Right)
            ]
            @ self.Left.ToGraphVizEdges
            @ self.Right.ToGraphVizEdges
        override self.ToGraphViz: string =
            let edges = self.ToGraphVizEdges
            let nums = new Dict<CRTNode<'a>,int>()
            let nodes = List.map (fun edge ->
                            match edge with
                            | Left(r,c) -> [r;c]
                            | Right(r,c) -> [r;c]
                        ) edges |> List.concat |> List.distinct
            let mutable i = 0
            for node in nodes do
                nums.Add(node, i)
                i <- i + 1

            let gznodes = List.map (fun (node: CRTNode<'a>) ->
                                match node with
                                | :? CRTRoot<'a> as r ->
                                    nums.[node].ToString() + " [label=\"root\"]"
                                | :? CRTLeaf<'a> as  l ->
                                    let prefix = l.Prefix.MaskedBitsAsString(UInt128.MaxValue)
                                    let value = l.Value.Value.ToString()
                                    nums.[node].ToString() + "[shape=record, label=\"{" + prefix + "|" + value + "}\"]"
                                | :? CRTEmptyLeaf<'a> as e ->
                                    let prefix = e.Prefix.MaskedBitsAsString(UInt128.MaxValue)
                                    let value = "ε"
                                    nums.[node].ToString() + "[shape=record, label=\"{" + prefix + "|" + value + "}\"]"
                                | :? CRTInner<'a> as i ->
                                    nums.[node].ToString() + " [label=\"" + i.Prefix.MaskedBitsAsString(i.Mask) + "\"]"
                          ) nodes
            let gzedges = List.map (fun edge ->
                              match edge with
                              | Left(r,c) -> nums.[r].ToString() + " -- " + nums.[c].ToString() + " [label=\"0\"]"
                              | Right(r,c) -> nums.[r].ToString() + " -- " + nums.[c].ToString() + " [label=\"1\"]"
                          ) edges

            "graph {\n" +
            System.String.Join("\n",gznodes) +
            System.String.Join("\n",gzedges) +
            "}\n"

    /// <summary>
    /// An inner node of a compressed radix tree.
    /// </summary>
    /// <param name="endpos">
    /// The end index in the UInt128 of this node's bitmask.
    /// </param>
    /// <param name="prefix">
    /// The UInt128 key associated with this node. The prefix bits
    /// used in comparisons are the node's bitmask AND'ed with
    /// the prefix.
    /// </param
    /// <param name="left">
    /// The left subtree.
    /// </param>
    /// <param name="right">
    /// The right subtree.
    /// </param>
    and CRTInner<'a when 'a : equality>(endpos: int, prefix: UInt128, left: CRTNode<'a>, right: CRTNode<'a>) =
        inherit CRTNode<'a>(endpos, prefix)
        let mymask = UInt128.calcMask 0 endpos
        let mybits = mymask.BitwiseAnd prefix
        let nextBitMask = UInt128.calcMask (endpos + 1) (endpos + 1)
        member self.Left = left
        member self.Right = right
        member self.PrefixLength = endpos + 1
        member self.Mask : UInt128 = mymask
        override self.Prefix = prefix
        override self.IsRoot = false
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
        override self.Delete(key: UInt128) : CRTNode<'a> =
            let keybits = mymask.BitwiseAnd key
            if mybits = keybits then
                let nextbit = nextBitMask.BitwiseAnd key
                if nextBitMask.GreaterThan nextbit then
                    let nleft = left.Delete key
                    match nleft with
                    | :? CRTEmptyLeaf<'a> ->
                        // this node is redundant; return the right subtree
                        right
                    | _ -> CRTInner(endpos, prefix, nleft, right) :> CRTNode<'a>
                else
                    let nright = right.Delete key
                    match nright with
                    | :? CRTEmptyLeaf<'a> ->
                        // this node is redundant; return the left subtree
                        left
                    | _ -> CRTInner(endpos, prefix, left, nright) :> CRTNode<'a>
            else
            // the user wants to delete a node not in the tree
            // return the tree unmodified
                self :> CRTNode<'a>
        override self.InsertOr(key: UInt128)(value': 'a)(keyexists: 'a -> 'a -> 'a) : CRTNode<'a> =
            let keybits = mymask.BitwiseAnd key
            if mybits = keybits then
                let nextbit = nextBitMask.BitwiseAnd key
                if nextBitMask.GreaterThan nextbit then
                    CRTInner(endpos, prefix, left.InsertOr key value' keyexists, right) :> CRTNode<'a>
                else
                    CRTInner(endpos, prefix, left, right.InsertOr key value' keyexists) :> CRTNode<'a>
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
                    CRTInner(pidx - 1, prefix', self, CRTLeaf(key, value')) :> CRTNode<'a>
                else
                    // current node goes on the right
                    CRTInner(pidx - 1, prefix', CRTLeaf(key, value'), self) :> CRTNode<'a>
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            self.InsertOr key value (fun i _ -> i)
        override self.EnumerateSubtree(key: UInt128)(value: UInt128) : seq<'a> =
            match self.LookupSubtree key value with
            | Some(st) -> st.LRTraversal
            | None -> Seq.empty
        override self.LRTraversal: seq<'a> =
            // inner nodes store no data
            Seq.concat [ self.Left.LRTraversal; self.Right.LRTraversal ]
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTInner<'a> as other ->
                let ok = self.Left = other.Left &&
                         self.Right = other.Right &&
                         self.PrefixLength = other.PrefixLength &&
                         self.Prefix = other.Prefix
                ok
            | _ -> false
        override self.GetHashCode() : int =
            self.Left.GetHashCode() ^^^ self.Right.GetHashCode()
        override self.ToGraphVizEdges : GZEdge<'a> list =
            [
                Left(self, self.Left);
                Right(self, self.Right)
            ]
            @ self.Left.ToGraphVizEdges
            @ self.Right.ToGraphVizEdges
        override self.ToGraphViz: string = failwith "no"

    and CRTLeaf<'a when 'a : equality>(prefix: UInt128, value: 'a) =
        inherit CRTNode<'a>(127, prefix)
        override self.Prefix = prefix
        override self.IsRoot = false
        override self.IsLeaf = true
        override self.IsEmpty = false
        override self.Value = Some value
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option = 
            Some (self :> CRTNode<'a>)
        override self.Lookup(str: UInt128) : 'a option = Some value
        override self.InsertOr(key: UInt128)(value': 'a)(keyexists: 'a -> 'a -> 'a) : CRTNode<'a> =
            // find longest common prefix
            let pidx = key.LongestCommonPrefix prefix
            if pidx = 128 then  // the keys are exactly the same
                CRTLeaf(prefix, keyexists value value') :> CRTNode<'a>
            else
                // insert a new parent
                let mask' = UInt128.calcMask 0 (pidx - 1)
                let prefix' = mask'.BitwiseAnd key

                // insert current subtree on the left or on the right of new parent node?
                let nextBitMask' = UInt128.calcMask pidx pidx
                let nextbit = nextBitMask'.BitwiseAnd prefix
                if nextBitMask'.GreaterThan nextbit then
                    // current node goes on the left
                    CRTInner(pidx - 1, prefix', self, CRTLeaf(key, value')) :> CRTNode<'a>
                else
                    // current node goes on the right
                    CRTInner(pidx - 1, prefix', CRTLeaf(key, value'), self) :> CRTNode<'a>
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            self.InsertOr key value (fun i _ -> i)
        override self.EnumerateSubtree(key: UInt128)(value: UInt128) : seq<'a> = self.LRTraversal
        override self.LRTraversal: seq<'a> = seq [ value ]
        override self.Delete(key: UInt128) : CRTNode<'a> =
            CRTEmptyLeaf(prefix) :> CRTNode<'a>
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTLeaf<'a> as other ->
                let ok = self.Prefix = other.Prefix &&
                         self.Value = other.Value
                ok
            | _ -> false
        override self.GetHashCode() : int =
            prefix.GetHashCode()
        override self.ToGraphVizEdges : GZEdge<'a> list = []
        override self.ToGraphViz: string = failwith "no"

    and CRTEmptyLeaf<'a when 'a : equality>(prefix: UInt128) =
        inherit CRTNode<'a>(127, prefix)
        override self.Prefix = prefix
        override self.IsRoot = false
        override self.IsLeaf = true
        override self.IsEmpty = true
        override self.Value = None
        override self.LookupSubtree(key: UInt128)(mask: UInt128) : CRTNode<'a> option = None
        override self.Lookup(str: UInt128) : 'a option = None
        override self.InsertOr(key: UInt128)(value': 'a)(keyexists: 'a -> 'a -> 'a) : CRTNode<'a> =
            CRTLeaf(key, value') :> CRTNode<'a>
        override self.Replace(key: UInt128)(value: 'a) : CRTNode<'a> =
            self.InsertOr key value (fun i _ -> i)
        override self.EnumerateSubtree(key: UInt128)(value: UInt128) : seq<'a> = self.LRTraversal
        override self.LRTraversal: seq<'a> = Seq.empty
        override self.Delete(key: UInt128) : CRTNode<'a> =
            // user wants to delete key not in tree;
            // do nothing
            self :> CRTNode<'a>
        override self.Equals(o: obj) : bool =
            match o with
            | :? CRTEmptyLeaf<'a> as other ->
                let ok = self.Prefix = other.Prefix
                ok
            | _ -> false
        override self.GetHashCode() : int =
            prefix.GetHashCode()
        override self.ToGraphVizEdges : GZEdge<'a> list = []
        override self.ToGraphViz: string = failwith "no"