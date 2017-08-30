using System;
using System.IO;
using CsvHelper;
using System.Collections.Generic;
using System.Linq;

namespace ExceLintFileFormats
{
    public class CorpusStats : IDisposable
    {
        StreamWriter _sw;
        CsvWriter _cw;
        HashSet<string> _processed = new HashSet<string>();

        public CorpusStats(string path)
        {
            // if CSV already exists, read it into memory
            // and append instead of creating the file
            var records = new List<CorpusStatsRow>();

            if (File.Exists(path))
            {
                // read old records
                using (var sr = new StreamReader(path))
                {
                    using (var cr = new CsvReader(sr))
                    {
                        var lastBuffer = new List<CorpusStatsRow>();

                        string last = null;
                        while (cr.Read())
                        {
                            var record = cr.GetRecord<CorpusStatsRow>();

                            // skip header
                            if (record.Workbook == "Workbook")
                            {
                                continue;
                            }

                            // if this record is for a different workbook
                            // than the last record, we know that the last
                            // record written was not interrupted
                            if (last != null && last != record.Workbook)
                            {
                                records.AddRange(lastBuffer);
                                lastBuffer.Clear();
                            }

                            // update last pointer
                            last = record.Workbook;

                            // add record to buffer
                            lastBuffer.Add(record);
                        }
                    }
                }

                // since the last record may have been interrupted
                // rewrite all records, and resume from there
                using(var sw = new StreamWriter(path, append: false))
                {
                    using (var cw = new CsvWriter(sw))
                    {
                        // write header
                        cw.WriteHeader<CorpusStatsRow>();

                        // write the rest
                        cw.WriteRecords(records);
                    }
                }

                // save where we left off
                _processed = new HashSet<string>(records.Select(row => row.Workbook));
            }

            // append if the CSV already contains something
            bool doAppend = _processed.Count > 0 ? true : false;
            _sw = new StreamWriter(path, append: doAppend);
            _sw.AutoFlush = true;
            _cw = new CsvWriter(_sw);

            // write header if not appending
            if (!doAppend)
            {
                _cw.WriteHeader<CorpusStatsRow>();
            }
        }

        public void WriteRow(CorpusStatsRow row)
        {
            _cw.WriteRecord(row);
        }

        public bool IsProcessed(string workbookFullPath)
        {
            return _processed.Contains(Path.GetFileName(workbookFullPath));
        }

        public bool WasResumed
        {
            get { return _processed.Count > 0; }
        }

        public void MarkAsProcessed(string workbookFullPath)
        {
            _processed.Add(Path.GetFileName(workbookFullPath));
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