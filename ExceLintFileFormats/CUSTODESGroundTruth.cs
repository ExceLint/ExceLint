using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExceLintFileFormats
{
    public class CUSTODESGroundTruth
    {
        CUSTODESGroundTruthRow[] _rows;

        private CUSTODESGroundTruth(CUSTODESGroundTruthRow[] rows)
        {
            _rows = rows;
        }

        public CUSTODESGroundTruthRow[] Rows
        {
            get { return _rows; }
        }

        public static CUSTODESGroundTruth Load(string path)
        {
            using (var sr = new StreamReader(path))
            {
                var rows = new CsvReader(sr).GetRecords<CUSTODESGroundTruthRow>().ToArray();

                return new CUSTODESGroundTruth(rows);
            }
        }
    }

    public class CUSTODESGroundTruthRow
    {
        public int Index { get; set; }
        public string Spreadsheet { get; set; }
        public string Worksheet { get; set; }
        public string GroundTruth { get; set; }
        public string CUSTODES { get; set; }
        public string AmCheck { get; set; }
        public string UCheck { get; set; }
        public string Dimension { get; set; }
        public string Excel { get; set; }
    }

    public sealed class CUSTODESGroundTruthRowMap : CsvClassMap<CUSTODESGroundTruthRow>
    {
        public CUSTODESGroundTruthRowMap()
        {
            Map(m => m.Index).Name("Index");
            Map(m => m.Spreadsheet).Name("Spreadsheet");
            Map(m => m.Worksheet).Name("Worksheet");
            Map(m => m.GroundTruth).Name("Ground truth");
            Map(m => m.CUSTODES).Name("Custodes");
            Map(m => m.AmCheck).Name("AmCheck");
            Map(m => m.UCheck).Name("UCheck");
            Map(m => m.Dimension).Name("Dimension");
            Map(m => m.Excel).Name("Excel");
        }
    }
}