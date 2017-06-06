using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using Cluster = System.Collections.Generic.HashSet<AST.Address>;
using Clustering = System.Collections.Generic.HashSet<System.Collections.Generic.HashSet<AST.Address>>;

namespace ExceLintUI
{
    public class ClusterColorer
    {
        // cluster neighbors
        Dictionary<Cluster, HashSet<Cluster>> cNeighbors;

        // address-to-cluster lookup
        Dictionary<AST.Address, Cluster> cdict;

        // color map
        Dictionary<Cluster, Color> assignedColors;

        public ClusterColorer(Clustering cs)
        {
            // init address-to-cluster lookup
            foreach(Cluster c in cs)
            {
                foreach (AST.Address addr in c)
                {
                    cdict.Add(addr, c);
                }
            }

            // init cluster neighbor map
            var cNeighbors = new Dictionary<HashSet<AST.Address>, HashSet<HashSet<AST.Address>>>();
            foreach (Cluster c in cs)
            {
                cNeighbors.Add(c, new HashSet<Cluster>());
                var neighbors = AdjacentCells(c);
                foreach (Cluster c2 in cs)
                {
                    // append if c is adjacent to c2
                    if (neighbors.Intersect(c2).Count() > 0)
                    {
                        cNeighbors[c].Add(c2);
                    }
                }
            }

            // rank clusters by their degree
            Cluster[] csSorted = cs.OrderByDescending(c => cNeighbors[c].Count).ToArray();
            
            // greedily assign colors by degree, largest first;
            // this is the Welsh-Powell heuristic
            foreach (Cluster c in csSorted)
            {
                // TODO
            }
        }

        public static IEnumerable<double> Angles(double start, double end)
        {
            var midpoint = (end - start) / 2 + start;
            yield return midpoint;

            // split this region into two regions, and recursively enumerate
            var top = Angles(start, midpoint);
            var bottom = Angles(midpoint, end);

            while (true)
            {
                yield return top.Take(1).First();
                top = top.Skip(1);

                yield return bottom.Take(1).First();
                bottom = bottom.Skip(1);
            }
        }

        /// <summary>
        /// Returns the set of cells in a bounding box around the given
        /// cluster.
        /// </summary>
        /// <param name="c">A cluster</param>
        /// <returns>The set of adjacent cells</returns>
        private static HashSet<AST.Address> AdjacentCells(Cluster c)
        {
            var hs = c.SelectMany(addr => AdjacentCells(addr))
                      .Aggregate(new HashSet<AST.Address>(), (acc, addr) => {
                          acc.Add(addr);
                          return acc;
                       });
            return hs;
        }

        /// <summary>
        /// Returns the set of cells adjacent to the given address.
        /// </summary>
        /// <param name="addr">An address</param>
        /// <returns>The set of adjacent cells</returns>
        private static HashSet<AST.Address> AdjacentCells(AST.Address addr)
        {
            var n  = Above(addr);
            var ne = Above(Right(addr));
            var e  = Right(addr);
            var se = Below(Right(addr));
            var s  = Below(addr);
            var sw = Below(Left(addr));
            var w  = Left(addr);
            var nw = Above(Left(addr));

            AST.Address[] addrs = { n, ne, e, se, s, sw, w, nw };

            return new HashSet<AST.Address>(addrs.Where(a => isSane(a)));
        }

        private static bool isSane(AST.Address addr)
        {
            return addr.Col > 0 && addr.Row > 0;
        }

        private static AST.Address Above(AST.Address addr)
        {
            return AST.Address.fromR1C1withMode(
                addr.Row - 1,
                addr.Col,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);
        }

        private static AST.Address Below(AST.Address addr)
        {
            return AST.Address.fromR1C1withMode(
                addr.Row + 1,
                addr.Col,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);
        }

        private static AST.Address Left(AST.Address addr)
        {
            return AST.Address.fromR1C1withMode(
                addr.Row,
                addr.Col - 1,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);
        }

        private static AST.Address Right(AST.Address addr)
        {
            return AST.Address.fromR1C1withMode(
                addr.Row,
                addr.Col + 1,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);
        }

        //public Color GetColor(Cluster c)
        //{

        //}
    }
}
