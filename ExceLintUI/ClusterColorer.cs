using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using static ExceLintUI.ColorCalc;
using Cluster = System.Collections.Generic.HashSet<AST.Address>;
using Clustering = System.Collections.Generic.HashSet<System.Collections.Generic.HashSet<AST.Address>>;
using System;

namespace ExceLintUI
{
    public class ClusterColorer
    {
        // fixed color attributes
        private static readonly double SATURATION = 1.0;
        private static readonly double LUMINOSITY = 0.5;

        // cluster neighbors
        Dictionary<Cluster, HashSet<Cluster>> cNeighbors;

        // address-to-cluster lookup
        Dictionary<AST.Address, Cluster> cdict;

        // color map
        Dictionary<Cluster, Color> assignedColors;

        public ClusterColorer(Clustering cs, double degreeStart, double degreeEnd)
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
            cNeighbors = new Dictionary<HashSet<AST.Address>, HashSet<HashSet<AST.Address>>>();
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
            // aka Welsh-Powell heuristic
            foreach (Cluster c in csSorted)
            {
                // get neighbor colors
                var ns = cNeighbors[c];
                var nscs = new HashSet<Color>();
                foreach (Cluster n in ns)
                {
                    if (assignedColors.ContainsKey(n))
                    {
                        nscs.Add(assignedColors[n]);
                    }
                }

                // init angle generator
                var angles = new AngleGenerator(degreeStart, degreeEnd);

                // color getter
                Func<Color> colorf = () => HSLtoColor(new HSL(angles.NextAngle(), SATURATION, LUMINOSITY));

                // get initial color
                var color = colorf();
                while(nscs.Contains(color))
                {
                    // get next color
                    color = colorf();
                }

                // save color
                assignedColors[c] = color;
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

        public Color GetColor(Cluster c)
        {
            return assignedColors[c];
        }
    }
}
