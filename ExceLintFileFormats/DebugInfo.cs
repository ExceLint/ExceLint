using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class DebugInfo : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public DebugInfo(string path)
        {
            _sw = new StreamWriter(path);
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<DebugInfoRow>();

            _sw.Flush();
        }

        public void WriteRow(DebugInfoRow row)
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

    public class DebugInfoRow
    {
        public string Path { get; set; }
        public string Workbook { get; set; }
        public string Worksheet { get; set; }
        public string Address { get; set; }
    }
}