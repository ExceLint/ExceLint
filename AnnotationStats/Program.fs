open System
open System.IO
open System.Collections.Generic
open ExceLintFileFormats
open CUSTODES

type BugClass = HashSet<AST.Address>

let count_true_ref_TP(etruth: ExceLintFileFormats.ExceLintGroundTruth)(flags: HashSet<AST.Address>) : int =
    let trueref = new Dictionary<BugClass*BugClass,int>()
        
    for addr in flags do
        // is it a bug?
        if etruth.IsAnInconsistentFormulaBug addr then
            // does it have a dual?
            if etruth.AddressHasADual addr then
                // get duals
                let duals = etruth.DualsForAddress addr
                // count
                if not (trueref.ContainsKey duals) then
                    trueref.Add(duals, 1)
                else
                    // add one if we have not exceeded our max count for this dual
                    if trueref.[duals] < (etruth.NumBugsForBugClass (fst duals)) then
                        trueref.[duals] <- trueref.[duals] + 1
            else
            // make a singleton bugclass and count it
                let bugclass = new HashSet<AST.Address>([addr])
                let duals = (bugclass,bugclass)
                trueref.Add(duals, 1)

    Seq.sum trueref.Values

[<EntryPoint>]
let main argv = 
    let trueref_file = argv.[0]
    let custodes_file = argv.[1]
    let custodes_wbdir = argv.[2]
            
    let excelint_gt = ExceLintGroundTruth.Load(trueref_file)
    let custodes_gt = GroundTruth.Load(custodes_wbdir, custodes_file)

    let gt_flags = custodes_gt.AllGroundTruth

    // get set of workbooks annotated in ExceLInt
    let wbs = new HashSet<string>(excelint_gt.WorkbooksAnnotated)

    // count the number of TP in hand-labeled CUSTODES ground truth file
    let custodes_num_TP = count_true_ref_TP excelint_gt gt_flags

    // count the total number of bugs of each class in true ref ground truth
    let excelint_num_inconsistent = excelint_gt.TotalNumBugKindBugs(ErrorClass.INCONSISTENT)
    let excelint_num_missing_formula = excelint_gt.TotalNumBugKindBugs(ErrorClass.MISSINGFORMULA)
    let excelint_num_whitespace = excelint_gt.TotalNumBugKindBugs(ErrorClass.WHITESPACE)
    let excelint_num_suspicious = excelint_gt.TotalNumBugKindBugs(ErrorClass.SUSPICIOUS)

    // count off-by-one bugs; requires dependence analysis
    let excelint_num_off_by_one = excelint_gt.TotalOffByOneBugs(custodes_wbdir);

    // count number of annotations in CUSTODES for workbooks in ExceLint
    let custodes_num_annot = (gt_flags |> Seq.filter (fun a -> wbs.Contains(a.WorkbookName)) |> Seq.length) 

    // count number of 'not a bug' from CUSTODES
    let num_custodes_notabug =
        gt_flags
        |> Seq.filter (fun a -> wbs.Contains a.WorkbookName)
        |> Seq.map (fun a -> excelint_gt.AnnotationFor a)
        |> Seq.filter (fun anno ->
               match anno.BugKind with
               | :? NotABug -> true
               | _ -> false
           )
        |> Seq.length

    // print stuff
    printfn "For the %A workbooks annotated in the ExceLint corpus:" wbs.Count
    printfn "There are %A inconsistent formula errors in the CUSTODES corpus." custodes_num_TP
    printfn "There are %A inconsistent formula errors in the ExceLint corpus." excelint_num_inconsistent
    printfn "There are %A more inconsistent formula errors in the ExceLint corpus than in the CUSTODES corpus." (excelint_num_inconsistent - custodes_num_TP)
    printfn "There are %A missing formulas in the ExceLint corpus" excelint_num_missing_formula
    printfn "There are %A whitespace bugs in the ExceLint corpus." excelint_num_whitespace
    printfn "There are %A suspicious cells in the ExceLint corpus." excelint_num_suspicious
    printfn "There are %A off-by-one errors in the ExceLint corpus." excelint_num_off_by_one
    printfn "There are %A total annotations for %A workbooks in the CUSTODES corpus." custodes_num_annot wbs.Count
    printfn "There are %A total annotations in the ExceLint corpus." excelint_gt.Flags.Count
    printfn "%A annotations from CUSTODES are labeled 'not a reference bug' in ExceLint corpus." num_custodes_notabug

    printfn "\nPress any key to continue."
    Console.ReadKey() |> ignore

    printfn "%A" argv
    0 // return an integer exit code
