module Args

open System.IO
open System.Text.RegularExpressions

    type Config(dpath: string, ofile: string, errfile: string) =
        do
            printfn "\n------------------------------------"
            printfn "Running with the following options: "
            printfn "------------------------------------"
            printfn "Input directory: %A" dpath
            printfn "Output file: %A" ofile
            printfn "Error file: %A" errfile
            printfn "------------------------------------\n"

        member self.files: string[] =
//            Directory.EnumerateFiles(dpath, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray
            Directory.EnumerateFiles(dpath, "*", SearchOption.AllDirectories) |> Seq.toArray
        member self.output_file: string = ofile
        member self.error_file: string = errfile 

    type Knobs = { verbose: bool; dont_exit: bool; alpha: double }

    let usage() : unit =
        printfn "StatsGatherer.exe <input directory> <output file.csv>"
        printfn 
            "Recursively finds all Excel (*.xls and *.xlsx) files in <input directory>, \n\
             opens them, and gathers a variety of stats on them, which are written to the \n\
             CSV file specified in <output file.csv>."

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length <> 2 || (Array.contains "--help" argv) then
            usage()
        let dpath  = System.IO.Path.GetFullPath argv.[0]   // input directory
        let opath  = System.IO.Path.GetFullPath argv.[1]   // output file

        let opath_dir = System.IO.Path.GetDirectoryName opath
        let opath_fne = System.IO.Path.GetFileNameWithoutExtension opath

        let errpath = Path.Combine(opath_dir, (opath_fne + "_err.csv"))  // error file

        Config(dpath, opath, errpath)

