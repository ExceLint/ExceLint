module PerfUtils

open System

type private TimeUnit =
| Milliseconds
| Ticks

let private run<'A,'B>(f: 'A -> 'B)(input: 'A)(tu: TimeUnit) : 'B*int64 =
    let _sw = new Diagnostics.Stopwatch()
    _sw.Start()
    let _o = f input
    _sw.Stop()
    let tu_amount =
        match tu with
        | Milliseconds -> _sw.ElapsedMilliseconds
        | Ticks -> _sw.ElapsedTicks
    (_o,tu_amount)

let runMillis<'A,'B>(f: 'A -> 'B)(input: 'A) : 'B*int64 =
    run f input Milliseconds

let runTicks<'A,'B>(f: 'A -> 'B)(input: 'A) : 'B*int64 =
    run f input Ticks