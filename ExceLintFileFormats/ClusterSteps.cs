using System;
using System.IO;
using CsvHelper;

namespace ExceLintFileFormats
{
    public class ClusterSteps : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;

        public ClusterSteps(string path)
        {
            _sw = new StreamWriter(path);
            _cw = new CsvWriter(_sw);

            // write header
            _cw.WriteHeader<ClusterStepsRow>();

            _sw.Flush();
        }

        public void WriteRow(ClusterStepsRow row)
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

    public class ClusterStepsRow
    {
        public bool Show { get; set; }
        public string Merge { get; set; }
        public double Distance { get; set; }
        public double FScore { get; set; }
        public double WCSS { get; set; }
        public double BCSS { get; set; }
        public double TSS { get; set; }
        public int k { get; set; }
    }
}