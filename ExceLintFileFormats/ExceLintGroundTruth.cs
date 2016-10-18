using System.IO;
using System.Collections.Generic;
using System.Linq;
using CsvHelper;




//        [< Literal >]

//        let DebugInfoHeaders = headers DebugInfoSchema

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
