using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class Clustering : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public Clustering(string path)
        {
            _sw = new StreamWriter(path);
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<ClusteringRow>();

            _sw.Flush();
        }

        public void WriteRow(ClusteringRow row)
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

    public class ClusteringRow
    {
        public string Path { get; set; }
        public string Workbook { get; set; }
        public string Worksheet { get; set; }
        public string Address { get; set; }
        public int Cluster { get; set; }
    }
}