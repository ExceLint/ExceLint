module CSV

open System.Text.RegularExpressions
open FSharp.Data

let fsdTypeRegex = new Regex(" \([a-z0-9]+\)", RegexOptions.Compiled)

[<Literal>]
let private ExceLintStatsSchema = "benchmark_name (string), num_cells (int), num_formulas (int),\
    sig_thresh (float), dep_time_ms (int64), score_time_ms (int64), freq_time_ms (int64), ranking_time_ms (int64),\
    causes_time_ms (int64), conditioning_set_sz_time_ms (int64), num_anom (int), custodes_fail (bool),\
    custodes_fail_msg (string), num_custodes_smells (int), true_smells (int),\
    excelint_true_smells_found (int), custodes_true_smells_found (int), excelint_custodes_true_smell_intersect (int),\
    true_smells_missed_by_both (int), excel_flags (int), excelint_excel_intersect (int), custodes_excel_intersect (int),\
    excel_flags_missed_by_both (int),\
    opt_spectral (bool),\
    opt_cond_all_cells (bool), opt_cond_rows (bool), opt_cond_cols (bool), opt_cond_levels (bool),opt_cond_sheets (bool),\
    opt_addrmode_inference (bool), opt_weight_intrinsic_anom (bool), opt_weight_condition_set_sz (bool)"

[<Literal>]
let private WorkbookStatsSchema = "path (string), workbook (string), worksheet (string), addr (string), is_formula (bool), flagged_by_excelint (bool), flagged_by_custodes (bool), flagged_by_excel (bool), cli_same_as_v1 (bool), rank (int), score (float), custodes_true_smell (bool)"

[<Literal>]
let private CUSTODESGroundTruthSchema = "Index (int),Spreadsheet (string), Worksheet (string), GroundTruth (string), Custodes (string), AmCheck (string), UCheck (string), Dimension (string), Excel (string)"

[<Literal>]
let private DebugInfoSchema = "path (string), workbook (string), worksheet (string), addr (string)"

let CUSTODESGroundTruthPath = "../../../../data/analyses/CUSTODES/smell_detection_result.csv"

type ExceLintStats = 
  CsvProvider<Schema = ExceLintStatsSchema, HasHeaders=false>

type WorkbookStats =
    CsvProvider<Schema = WorkbookStatsSchema, HasHeaders=false>

type CUSTODESGroundTruth = CsvProvider<Schema = CUSTODESGroundTruthSchema, HasHeaders=false>

type DebugInfo = CsvProvider<Schema = DebugInfoSchema, HasHeaders=false>

let private headers(schema: string) : string = (fsdTypeRegex.Replace(schema, "") + "\n").Replace(" ", "")

let ExceLintStatsHeaders = headers ExceLintStatsSchema

let WorkbookStatsHeaders = headers WorkbookStatsSchema

let DebugInfoHeaders = headers DebugInfoSchema