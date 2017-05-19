namespace ExceLint
open System.Collections.Generic
    module Utils =
        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>

        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)
        let argwhatever(f: 'a -> double)(xs: seq<'a>)(whatev: double -> double -> bool) : 'a =
                Seq.reduce (fun arg x -> if whatev (f arg) (f x) then arg else x) xs
        let argmax(f: 'a -> double)(xs: seq<'a>) : 'a =
            argwhatever f xs (fun a b -> a > b)
        let argmin(f: 'a -> double)(xs: seq<'a>) : 'a =
            argwhatever f xs (fun a b -> a < b)

        let argindexmin(f: 'a -> int)(xs: 'a[]) : int =
            let fx = Array.map f xs

            Array.mapi (fun i res -> (i, res)) fx |>
            Array.fold (fun arg (i, res) ->
                if arg = -1 || res < fx.[arg] then
                    i
                else
                    arg
            ) -1 

        let HsUnion<'a>(h1: HashSet<'a>)(h2: HashSet<'a>) : HashSet<'a> =
            let h3 = new HashSet<'a>(h1)
            h3.UnionWith h2
            h3