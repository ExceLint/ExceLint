using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class CorpusStats : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public CorpusStats(string path)
        {
            _sw = new StreamWriter(path);
            _sw.AutoFlush = true;
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<DebugInfoRow>();
        }

        public void WriteRow(CorpusStatsRow row)
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

    public class CorpusStatsRow
    {
        public string Workbook { get; set; }
        public string Variable { get; set; }
        public long Value { get; set; }
    }
}