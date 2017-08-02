using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Collections.Immutable;
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

        public static ClusteringRow[] clusteringToRows(HashSet<HashSet<AST.Address>> clustering, Dictionary<HashSet<AST.Address>, int> ids)
        {
            var rows = new LinkedList<ClusteringRow>();

            foreach (HashSet<AST.Address> cluster in clustering)
            {
                foreach (AST.Address addr in cluster)
                {
                    var row = new ClusteringRow();
                    row.Path = addr.A1Path();
                    row.Workbook = addr.A1Workbook();
                    row.Worksheet = addr.A1Worksheet();
                    row.Address = addr.A1Local();
                    row.Cluster = ids[cluster];
                    rows.AddLast(row);
                }
            }

            var sorted_rows = rows.OrderBy(row => new Tuple<string, string, string, string>(row.Path, row.Workbook, row.Worksheet, row.Address));
            return sorted_rows.ToArray();
        }

        public static void writeClustering(HashSet<HashSet<AST.Address>> clustering, Dictionary<HashSet<AST.Address>, int> ids, string filename)
        {
            var rows = clusteringToRows(clustering, ids);

            using(var csv = new Clustering(filename))
            {
                foreach (var row in rows)
                {
                    csv.WriteRow(row);
                }
            }
        }

        public static void writeClustering(ImmutableHashSet<ImmutableHashSet<AST.Address>> clustering, string filename)
        {
            var mclustering = new HashSet<HashSet<AST.Address>>(clustering.Select(cl => new HashSet<AST.Address>(cl)));
            writeClustering(mclustering, filename);
        }

        public static void writeClustering(HashSet<HashSet<AST.Address>> clustering, string filename)
        {
            int i = 0;
            var ids = new Dictionary<HashSet<AST.Address>, int>();
            foreach(var cl in clustering)
            {
                ids.Add(cl, i);
                i++;
            }

            var rows = clusteringToRows(clustering, ids);

            using (var csv = new Clustering(filename))
            {
                foreach (var row in rows)
                {
                    csv.WriteRow(row);
                }
            }
        }

        public static ClusteringRow[] readClusteringAsRows(string filename)
        {
            ClusteringRow[] records;
            using (TextReader tw = File.OpenText(filename))
            {
                var csv = new CsvReader(tw);
                records = csv.GetRecords<ClusteringRow>().ToArray();
            }
            
            return records;
        }

        public static HashSet<HashSet<AST.Address>> readClustering(string filename)
        {
            var rows = readClusteringAsRows(filename);

            var clustering = new HashSet<HashSet<AST.Address>>();
            var ids = new Dictionary<int, HashSet<AST.Address>>();
            foreach (var row in rows)
            {
                if (!ids.ContainsKey(row.Cluster))
                {
                    var hs = new HashSet<AST.Address>();
                    ids.Add(row.Cluster, hs);
                    clustering.Add(hs);
                }

                var addr = AST.Address.FromA1StringForceMode(
                            row.Address,
                            AST.AddressMode.Absolute,
                            AST.AddressMode.Absolute,
                            row.Worksheet,
                            row.Workbook,
                            row.Path
                           );
                ids[row.Cluster].Add(addr);
            }

            return clustering;
        }
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