module CSV

open System.Text.RegularExpressions
open FSharp.Data

let fsdTypeRegex = new Regex(" \([a-z0-9]+\)", RegexOptions.Compiled)

[<Literal>]
let private ExceLintStatsSchema = "benchmark_name (string), num_cells (int), num_formulas (int),\
    sig_thresh (float), dep_time_ms (int64), score_time_ms (int64), freq_time_ms (int64), ranking_time_ms (int64),\
    causes_time_ms (int64), conditioning_set_sz_time_ms (int64), num_anom (int), num_custodes_smells (int),\
    opt_cond_all_cells (bool), opt_cond_rows (bool), opt_cond_cols (bool), opt_cond_levels (bool),\
    opt_addrmode_inference (bool), opt_weight_intrinsic_anom (bool), opt_weight_condition_set_sz (bool)"

[<Literal>]
let private WorkbookStatsSchema = "flagged_cell_addr (string), flagged_by_excelint (bool), flagged_by_custodes (bool), cli_same_as_v1 (bool), rank (int), score (float), custodes_true_smell (bool)"

[<Literal>]
let CUSTODESGroundTruthSchema = "Index (int),Spreadsheet (string), Worksheet (string), GroundTruth (string), Custodes (string), AmCheck (string), UCheck (string), Dimension (string), Excel (string)"

let CUSTODESGroundTruthPath = "../../../../data/analyses/CUSTODES/smell_detection_result.csv"

type ExceLintStats = 
  CsvProvider<Schema = ExceLintStatsSchema, HasHeaders=false>

type WorkbookStats =
    CsvProvider<Schema = WorkbookStatsSchema, HasHeaders=false>

type CUSTODESGroundTruth = CsvProvider<Schema = CUSTODESGroundTruthSchema, HasHeaders=false>

let ExceLintStatsHeaders = fsdTypeRegex.Replace(ExceLintStatsSchema, "") + "\n"

let WorkbookStatsHeaders = fsdTypeRegex.Replace(WorkbookStatsSchema, "") + "\n"