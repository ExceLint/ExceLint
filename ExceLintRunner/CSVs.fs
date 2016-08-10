module CSV

open FSharp.Data

type ExceLintStats = 
  CsvProvider<Schema = "benchmark_name (string), num_cells (int), num_formulas (int),\
    sig_thresh (float), dep_time_ms (int64), score_time_ms (int64), freq_time_ms (int64), ranking_time_ms (int64),\
    causes_time_ms (int64), conditioning_set_sz_time_ms (int64), num_anom (int), num_custodes_smells (int),\
    opt_cond_all_cells (bool), opt_cond_rows (bool), opt_cond_cols (bool), opt_cond_levels (bool),\
    opt_addrmode_inference (bool), opt_weight_intrinsic_anom (bool), opt_weight_condition_set_sz (bool)", HasHeaders=false>

type WorkbookStats =
    CsvProvider<Schema = "flagged_cell_addr (string), flagged_by_excelint (bool), flagged_by_custodes (bool), rank (int), score (float), custodes_true_smell (bool)", HasHeaders=false>

let ExceLintStatsHeaders = "benchmark_name,num_cells,num_formulas,sig_thresh,dep_time_ms,score_time_ms,\
freq_time_ms,ranking_time_ms,causes_time_ms,conditioning_set_sz_time_ms,num_anom,num_custodes_smells,opt_cond_all_cells,\
opt_cond_rows,opt_cond_cols,opt_cond_levels,opt_addrmode_inference,opt_weight_intrinsic_anom,\
opt_weight_condition_set_sz\n"