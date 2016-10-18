using System.IO;
using System.Collections.Generic;
using System.Linq;
using CsvHelper;

//namespace ExceLint

//    open System.Text.RegularExpressions
//    open FSharp.Data

//    module CSV =
//        let fsdTypeRegex = new Regex(" \([a-z0-9]+\)", RegexOptions.Compiled)

//        [< Literal >]
//        let private ExceLintStatsSchema = "benchmark_name (string), num_cells (int), num_formulas (int),\
//            sig_thresh (float), dep_time_ms (int64), score_time_ms (int64), freq_time_ms (int64), ranking_time_ms (int64),\
//            causes_time_ms (int64), conditioning_set_sz_time_ms (int64), excelint_flags (int), min_anom_score (float), custodes_fail (bool),\
//            custodes_fail_msg (string),\
//            excelint_true_ref_bugs_found (int), custodes_true_ref_bugs_found (int),\
//            num_custodes_smells (int), true_smells (int),\
//            excelint_true_smells_found (int), custodes_true_smells_found (int), excelint_custodes_true_smell_intersect (int),\
//            true_smells_missed_by_both (int), excel_flags (int), excelint_excel_intersect (int), custodes_excel_intersect (int),\
//            excel_flags_missed_by_both (int),\
//            opt_spectral (bool),\
//            opt_cond_all_cells (bool), opt_cond_rows (bool), opt_cond_cols (bool), opt_cond_levels (bool),opt_cond_sheets (bool),\
//            opt_addrmode_inference (bool), opt_weight_intrinsic_anom (bool), opt_weight_condition_set_sz (bool)"

//        [<Literal>]
//let private WorkbookStatsSchema = "path (string), workbook (string), worksheet (string), addr (string), is_formula (bool), flagged_by_excelint (bool), flagged_by_custodes (bool), flagged_by_excel (bool), cli_same_as_v1 (bool), rank (int), score (float), excelint_true_bug (bool), custodes_true_smell (bool)"

//        [<Literal>]
//let private CUSTODESGroundTruthSchema = "Index (int),Spreadsheet (string), Worksheet (string), GroundTruth (string), Custodes (string), AmCheck (string), UCheck (string), Dimension (string), Excel (string)"

//        [<Literal>]
//let private OurGroundTruthSchema = "path (string), workbook (string), worksheet (string), addr (string), bug_kind (string), notes (string)"

//        // I need this in a static context and since F# has no facility
//        // for template metaprogramming, I just have to do it by hand...
//        [<Literal>]
//let OurGroundTruthHeaders = "path,workbook,worksheet,addr,bug_kind,notes"

//        [< Literal >]
//        let private DebugInfoSchema = "path (string), workbook (string), worksheet (string), addr (string)"

//        let CUSTODESGroundTruthPath = "../../../../data/analyses/CUSTODES/smell_detection_result.csv"

//        let private headers(schema: string) : string = (fsdTypeRegex.Replace(schema, "") + "\n").Replace(" ", "")

//        let ExceLintStatsHeaders = headers ExceLintStatsSchema

//        let WorkbookStatsHeaders = headers WorkbookStatsSchema

//        let DebugInfoHeaders = headers DebugInfoSchema

//        type ExceLintStats = 
//          CsvProvider<Schema = ExceLintStatsSchema, HasHeaders=false>

//        type WorkbookStats =
//            CsvProvider<Schema = WorkbookStatsSchema, HasHeaders=false>

//        type CUSTODESGroundTruth = CsvProvider<Schema = CUSTODESGroundTruthSchema, HasHeaders=false>

//        type OurGroundTruth = CsvProvider<Schema = OurGroundTruthSchema, HasHeaders = true, Sample = OurGroundTruthHeaders>

//        type DebugInfo = CsvProvider<Schema = DebugInfoSchema, HasHeaders=false>


namespace ExceLintFileFormats
{
    class ExceLintGroundTruth
    {
        public Dictionary<AST.Address, BugKind> _bugs = new Dictionary<AST.Address, BugKind>();
        public Dictionary<AST.Address, string> _notes = new Dictionary<AST.Address, string>();

        private AST.Address Address(string addrStr, string worksheetName, string workbookName, string path)
        {
            // we force the mode to absolute because
            // that's how Depends reads them
            return AST.Address.FromA1StringForceMode(
                addrStr.ToUpper(),
                AST.AddressMode.Absolute,
                AST.AddressMode.Absolute,
                worksheetName,
                (workbookName.EndsWith(".xls") ? workbookName : workbookName + ".xls"),
                Path.GetFullPath(path)   // ensure absolute path
            );
        }

        private ExceLintGroundTruth(ExceLintGroundTruthRow[] rows)
        {
            foreach (var row in rows)
            {
                AST.Address addr = Address(row.Address, row.Worksheet, row.Workbook, row.Path);
                _bugs.Add(addr, BugKind.ToKind(row.BugKind));
                _notes.Add(addr, row.Notes);
            }
        }

        public static ExceLintGroundTruth Load(string gtpath)
        {
            using (var sr = new StreamReader(gtpath))
            {
                var rows = new CsvReader(sr).GetRecords<ExceLintGroundTruthRow>().ToArray();

                return new ExceLintGroundTruth(rows);
            }
        }

        public static ExceLintGroundTruth Create(string gtpath)
        {
            using (StreamWriter sw = new StreamWriter(gtpath))
            {
                using (CsvWriter cw = new CsvWriter(sw))
                {
                    cw.WriteHeader<ExceLintGroundTruthRow>();
                }
            }

            return Load(gtpath);
        }
    }

    class ExceLintGroundTruthRow
    {
        public string Path { get; set; }
        public string Workbook { get; set; }
        public string Worksheet { get; set; }
        public string Address { get; set; }
        public string BugKind { get; set; }
        public string Notes { get; set; }
    }
}
