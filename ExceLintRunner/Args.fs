module Args

open System.IO

    type Config(dpath: string, csv: string) =
        member self.files: string[] =
            Directory.EnumerateFiles(dpath, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray
        member self.csv: string = csv

    let usage() : unit =
        printfn "ExceLintRunner [input directory] [output csv]"
        printfn (
            "Recursively finds all Excel (*.xls and *.xlsx) files in [input directory], \
            opens them, runs ExceLint, and prints output statistics to [output csv]. \
            Press Ctrl-C to cancel an analysis."
        )

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length <> 2 then
            usage()
        let dpath = argv.[0]
        let csv = argv.[1]
        Config(dpath, csv)

