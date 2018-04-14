using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class ExceLintStats : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public ExceLintStats(string path)
        {
            // create directory unless it already exists
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            _sw = new StreamWriter(path);
            _sw.AutoFlush = true;
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<ExceLintStatsRow>();
        }

        public void WriteRow(ExceLintStatsRow row)
        {
            _cw.WriteRecord(row);
        }

        #region IDisposable Support
        // To detect redundant calls
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _sw.Flush();
                    _cw.Dispose();

                    _cw = null;
                    _sw = null;
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }

    public class ExceLintStatsRow
    {
        public string BenchmarkName { get; set; }
        public int NumCells { get; set; }
        public int NumFormulas { get; set; }
        public double SigThresh { get; set; }
        public long TimeMarshalingMs { get; set; }
        public long TimeParsingMs { get; set; }
        public long TimeGraphConstruct { get; set; }
        public long ScoreTimeMs { get; set; }
        public long FreqTimeMs { get; set; }
        public long RankingTimeMs { get; set; }
        public long CausesTimeMs { get; set; }
        public long ConditioningSetSzTimeMs { get; set; }
        public int ExceLintFlags { get; set; }
        public double ExceLintPrecisionVsCustodesGT { get; set; }
        public double ExceLintRecallVsCustodesGT { get; set; }
        public double CUSTODESPrecisionVsCustodesGT { get; set; }
        public double CUSTODESRecallVsCustodesGT { get; set; }
        public int NumCUSTODESSmells { get; set; }
        public int NumTrueSmells { get; set; }
        public int NumExceLintTrueSmellsFound { get; set; }
        public int NumCUSTODESTrueSmellsFound { get; set; }
        public double MinAnomScore { get; set; }
        public long CUSTODESTimeMs { get; set; }
        public bool CUSTODESFailed { get; set; }
        public string CUSTODESFailureMsg { get; set; }
        public int NumTrueRefBugs { get; set; }
        public int ExceLintTrueRefTruePositives { get; set; }
        public int ExceLintTrueRefFalsePositives { get; set; }
        public int ExceLintMissingFormulaTruePositives { get; set; }
        public int ExceLintMissingFormulaFalsePositives { get; set; }
        public int ExceLintWhitespaceOpTruePositives { get; set; }
        public int ExceLintWhitespaceOpFalsePositives { get; set; }
        public double ExceLintPrecisionVsTrueRefBugs { get; set; }
        public double ExceLintRecallVsTrueRefBugs { get; set; }
        public double ExceLintTrueRefPValue { get; set; }
        public double ExceLintRandomTPBaseline { get; set; }
        public int CUSTODESTrueRefTruePositives { get; set; }
        public int CUSTODESTrueRefFalsePositives { get; set; }
        public int CUSTODESMissingFormulaTruePositives { get; set; }
        public int CUSTODESMissingFormulaFalsePositives { get; set; }
        public int CUSTODESWhitespaceOpTruePositives { get; set; }
        public int CUSTODESWhitespaceOpFalsePositives { get; set; }
        public double CUSTODESPrecisionVsTrueRefBugs { get; set; }
        public double CUSTODESRecallVsTrueRefBugs { get; set; }
        public double CUSTODESTrueRefPValue { get; set; }
        public double CUSTODESRandomTPBaseline { get; set; }
        public int NumExceLintCUSTODESTrueSmellsIntersect { get; set; }
        public int NumTrueSmellsMissedByBoth { get; set; }
        public int NumExcelFlags { get; set; }
        public int NumExceLintExcelIntersect { get; set; }
        public int NumCUSTODESExcelIntersect { get; set; }
        public int NumExcelMissedByBoth { get; set; }
        public bool OptSpectral { get; set; }
        public bool OptCondAllCells { get; set; }
        public bool OptCondRows { get; set; }
        public bool OptCondCols { get; set; }
        public bool OptCondLevels { get; set; }
        public bool OptCondSheets { get; set; }
        public bool OptAddrModeInference { get; set; }
        public bool OptWeightIntrinsicAnom { get; set; }
        public bool OptWeightConditionSetSz { get; set; }
        public double ExceLintJaccardDistance { get; set; }
        public int ExceLintDeltaK { get; set; }
        public int Collisions { get; set; }
    }
}
