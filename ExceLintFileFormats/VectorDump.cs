using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class VectorDump : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public VectorDump(string path)
        {
            _sw = new StreamWriter(path);
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<VectorDumpRow>();

            _sw.Flush();
        }

        public void WriteRow(VectorDumpRow row)
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

    public class VectorDumpRow
    {
        public int clusterID { get; set; }
        public double x { get; set; }
        public double y { get; set; }
        public double z { get; set; }
        public double dx { get; set; }
        public double dy { get; set; }
        public double dz { get; set; }
        public double dc { get; set; }
    }
}