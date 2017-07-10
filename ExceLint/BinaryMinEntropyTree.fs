namespace ExceLint

    open System.Collections.Generic
    open CommonTypes
    open Utils

    type Cells = Dict<AST.Address,Countable>
    type Coord = AST.Address * AST.Address
    type LeftTop = Coord
    type RightBottom = Coord
    type Region = LeftTop * RightBottom

    [<AbstractClass>]
    type BinaryMinEntropyTree() =
        static member private AddressEntropy(addrs: AST.Address[])(rmap: Dict<AST.Address,Countable>) : double =
            if addrs.Length = 0 then
                // the decomposition where one side is the empty set
                // is the worst decomposition, unless it is the only
                // decomposition
                System.Double.PositiveInfinity
            else
                // get values
                let vs = addrs |> Array.map (fun a -> rmap.[a].ToCVectorResultant)

                // count
                let cs = BasicStats.counts vs

                // compute probability vector
                let ps = BasicStats.empiricalProbabilities cs vs.Length

                // debug string
                let debug = "x = c(" + System.String.Join(", ", cs) + ")"

                // compute entropy
                let entropy = BasicStats.entropy ps

                entropy

        static member private MinEntropyPartition(rmap: Cells)(vert: bool) : AST.Address[]*AST.Address[] =
            // which axis we use depends on whether we are
            // decomposing verticalls or horizontally
            let indexer = (fun (a: AST.Address) -> if vert then a.X else a.Y)

            // extract addresses
            let addrs = Array.ofSeq rmap.Keys

            // find bounding box
            let (lt,rb) = Utils.BoundingRegion (rmap.Keys) 0

            let parts = Array.map (fun i ->
                               // partition addresses by "less than index of a",
                               // e.g., if vert then "less than x"
                               let (left,right) = addrs |> Array.partition (fun a -> indexer a < i)
                               left, right
                           ) [| indexer lt .. indexer rb + 1 |]

            let debug_es =
                parts
                |> Array.map (fun (l,r) ->
                    // compute entropy
                    let entropy_left = BinaryMinEntropyTree.AddressEntropy l rmap
                    let entropy_right = BinaryMinEntropyTree.AddressEntropy r rmap

                    // total for left and right
                    (entropy_left + entropy_right),entropy_left,entropy_right
                )

            parts
            |> Utils.argmin (fun (l,r) ->
                // compute entropy
                let entropy_left = BinaryMinEntropyTree.AddressEntropy l rmap
                let entropy_right = BinaryMinEntropyTree.AddressEntropy r rmap

                // total for left and right
                entropy_left + entropy_right
            )


        static member Infer(addrs: AST.Address[])(hb_inv: InvertedHistogram) : BinaryMinEntropyTree =
            let d = new Dict<AST.Address, Countable>()
            for addr in addrs do
                let (_,_,c) = hb_inv.[addr]
                d.Add(addr, c)
            BinaryMinEntropyTree.Decompose d None

        static member private Decompose(rmap: Cells)(parent_opt: Inner option) : BinaryMinEntropyTree =
            // find the minimum entropy decompositions
            let (left,right) = BinaryMinEntropyTree.MinEntropyPartition rmap true
            let (top,bottom) = BinaryMinEntropyTree.MinEntropyPartition rmap false

            // compute entropies again
            let e_vert = (BinaryMinEntropyTree.AddressEntropy left rmap) +
                         (BinaryMinEntropyTree.AddressEntropy right rmap)
            let e_horz = (BinaryMinEntropyTree.AddressEntropy top rmap) +
                         (BinaryMinEntropyTree.AddressEntropy bottom rmap)

            // split vertically or horizontally (favor vert for ties)
            if e_vert <= e_horz then
                // but is e_vert smaller because e_horz is a perfect decomposition?
                if e_horz = System.Double.PositiveInfinity then
                    Leaf(parent_opt, rmap) :> BinaryMinEntropyTree
                else
                    let l_rmap = left   |> Array.map (fun a -> a,rmap.[a]) |> adict
                    let r_rmap = right  |> Array.map (fun a -> a,rmap.[a]) |> adict

                    let node = Inner()
                    let lnode = BinaryMinEntropyTree.Decompose l_rmap (Some node)
                    let rnode = BinaryMinEntropyTree.Decompose r_rmap (Some node)
                    node.AddLeft lnode
                    node.AddRight rnode
                    node :> BinaryMinEntropyTree
            else
                // but is e_horz smaller because e_vert is a perfect decomposition?
                if e_vert = System.Double.PositiveInfinity then
                    Leaf(parent_opt, rmap) :> BinaryMinEntropyTree
                else
                    let t_rmap = top    |> Array.map (fun a -> a,rmap.[a]) |> adict
                    let b_rmap = bottom |> Array.map (fun a -> a,rmap.[a]) |> adict

                    let node = Inner()
                    let tnode = BinaryMinEntropyTree.Decompose t_rmap (Some node)
                    let bnode = BinaryMinEntropyTree.Decompose b_rmap (Some node)
                    node.AddLeft tnode
                    node.AddRight bnode
                    node :> BinaryMinEntropyTree

        /// <summary>return the leaves of the tree, in order of smallest to largest region</summary>
        static member Regions(tree: BinaryMinEntropyTree) : Leaf[] =
            match tree with
            | :? Inner as i -> Array.append (BinaryMinEntropyTree.Regions (i.Left)) (BinaryMinEntropyTree.Regions (i.Right))
            | :? Leaf as l -> [| l |]
            | _ -> failwith "Unknown tree node type."

    and Inner() =
        inherit BinaryMinEntropyTree()
        let mutable left = None
        let mutable right = None
        let mutable parent = None
        member self.AddParent(p: Inner) : unit =
            parent <- Some p
        member self.AddLeft(n: BinaryMinEntropyTree) : unit =
            left <- Some n
        member self.AddRight(n: BinaryMinEntropyTree) : unit =
            right <- Some n
        member self.Left =
            match left with
            | Some l -> l
            | None -> failwith "Cannot traverse tree until it is constructed!"
        member self.Right =
            match right with
            | Some r -> r
            | None -> failwith "Cannot traverse tree until it is constructed!"

    and Leaf(parent: Inner option, cells: Cells) =
        inherit BinaryMinEntropyTree()
