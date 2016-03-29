namespace PerfUtils

open System

type TimedLambda<'A,'B>(f: 'A -> 'B, input: 'A) = 
    let _sw = new Diagnostics.Stopwatch()
    let mutable _o : 'B option = None
    do
        _sw.Start()
        _o <- Some (f input)
        _sw.Stop()
    member self.Output =
        match _o with
        | Some(o) -> o
        | None -> failwith "Something went wrong!"
    member self.ElapsedMilliseconds = _sw.ElapsedMilliseconds
    member self.ElapsedTicks = _sw.ElapsedTicks