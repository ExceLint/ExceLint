using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class ExceptionLog : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public ExceptionLog(string path, bool append)
        {
            _sw = new StreamWriter(path, append);
            _sw.AutoFlush = true;
            _cw = new CsvWriter(_sw);

            // write header
            if (!append)
            {
                _cw.WriteHeader<ExceptionLogRow>();
            }
        }

        public void WriteRow(ExceptionLogRow row)
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

    public class ExceptionLogRow
    {
        public string Workbook { get; set; }
        public string Error { get; set; }
    }
}