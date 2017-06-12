module InstallScript

open System.IO
open System.Reflection

let DIR_ROOT = "custodes-bin"
let DIR_LIB  = "cc2_lib"
let DIR_OUTPUT = "output"
let DIR_INPUT = "input"
let JAR_CUSTODES = "cc2.jar"

let InitDirs() =
    // root path
    let root = Path.Combine(Path.GetTempPath(), DIR_ROOT) 
    Directory.CreateDirectory(root) |> ignore

    // libs
    let cc2_lib = Path.Combine(root, DIR_LIB)
    Directory.CreateDirectory(cc2_lib) |> ignore

    // output
    let output = Path.Combine(root, DIR_OUTPUT)
    Directory.CreateDirectory(output) |> ignore

    // input
    let input = Path.Combine(root, DIR_INPUT)
    Directory.CreateDirectory(input) |> ignore

    root

let AppJARPath(path: string) =
    Path.Combine(path, JAR_CUSTODES)

let TempSpreadsheetDir(path: string) =
    Path.Combine(path, DIR_INPUT)

let WriteResourceBinaryFiles(rootPath: string) =
    let assembly = Assembly.GetExecutingAssembly()
    let embeddedFiles = assembly.GetManifestResourceNames()

    for file in embeddedFiles do
        // only put cc2.jar in the top-level directory
        let dst = if file = JAR_CUSTODES then rootPath else Path.Combine(rootPath, DIR_LIB)

        // skip non-JARs
        if file.EndsWith(".jar") then
            using (assembly.GetManifestResourceStream(file)) (fun rsrc ->
                let newfile = Path.Combine(dst, file)
                using (new FileStream(newfile, FileMode.Create)) (fun outfile ->
                    rsrc.CopyTo(outfile)
                )
            )

let InstallCUSTODES(path: string) : string =
    WriteResourceBinaryFiles(path)
    AppJARPath(path)