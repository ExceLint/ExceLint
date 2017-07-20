namespace ExceLint
    open System.Collections.Generic

    module HashSetUtils =
        let difference<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.ExceptWith(hs2)
            hs3

        let differenceElem<'a>(hs: HashSet<'a>)(elem: 'a) : HashSet<'a> =
            let hs2 = new HashSet<'a>(hs)
            hs2.Remove(elem) |> ignore
            hs2

        let intersection<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.IntersectWith(hs2)
            hs3

        let union<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.UnionWith(hs2)
            hs3

        let unionElem<'a>(hs: HashSet<'a>)(elem: 'a) : HashSet<'a> =
            let hs2 = new HashSet<'a>(hs)
            hs2.Add elem |> ignore
            hs2

        let inPlaceUnion<'a>(source: HashSet<'a>)(target: HashSet<'a>) : unit =
            target.UnionWith(source)

        let inPlaceUnionElem<'a>(hs: HashSet<'a>)(elem: 'a) : unit =
            hs.Add elem |> ignore

        let equals<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : bool =
            let hsu = union hs1 hs2
            hs1.Count = hsu.Count