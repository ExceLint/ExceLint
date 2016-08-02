module CSV

open FSharp.Data

type ExceLintStats = 
  CsvProvider<Schema = "benchmark_name (string), num_cells (int), num_formulas (int),\
    sig_thresh (float), dep_time_ms (int64), score_time_ms (int64), freq_time_ms (int64), ranking_time_ms (int64),\
    causes_time_ms (int64), conditioning_set_sz_time_ms (int64), num_anom (int),\
    opt_cond_all_cells (bool), opt_cond_rows (bool), opt_cond_cols (bool), opt_cond_levels (bool),\
    opt_addrmode_inference (bool), opt_weight_intrinsic_anom (bool), opt_weight_condition_set_sz (bool)", HasHeaders=false>

type WorkbookStats =
    CsvProvider<Schema = "flaggedCellAddr (string), rank (int), score (float)", HasHeaders=false>