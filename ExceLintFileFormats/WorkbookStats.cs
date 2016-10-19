using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class WorkbookStats : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public WorkbookStats(string path)
        {
            _sw = new StreamWriter(path);
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<WorkbookStatsRow>();

            _sw.Flush();
        }

        public void WriteRow(WorkbookStatsRow row)
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

    public class WorkbookStatsRow
    {
        public string Path { get; set; }
        public string Workbook { get; set; }
        public string Worksheet { get; set; }
        public string Address { get; set; }
        public bool IsFormula { get; set; }
        public bool IsFlaggedByExceLint { get; set; }
        public bool IsFlaggedByCUSTODES { get; set; }
        public bool IsFlaggedByExcel { get; set; }
        public bool CLISameAsV1 { get; set; }
        public int Rank { get; set; }
        public double Score { get; set; }
        public bool IsExceLintTrueBug { get; set; }
        public bool IsCUSTODESTrueSmell { get; set; }
    }
}