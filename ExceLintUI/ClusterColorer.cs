using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using static ExceLintUI.ColorCalc;
using Cluster = System.Collections.Generic.HashSet<AST.Address>;
using Clustering = System.Collections.Generic.HashSet<System.Collections.Generic.HashSet<AST.Address>>;
using ROInvertedHistogram = System.Collections.Immutable.ImmutableDictionary<AST.Address, System.Tuple<string, ExceLint.Scope.SelectID, ExceLint.Countable>>;
using Fingerprint = ExceLint.Countable;
using System;
using FastDependenceAnalysis;

namespace ExceLintUI
{
    public class ClusterColorer
    {
        // fixed color attributes
        private static readonly double SATURATION = 1.0;
        private static readonly double LUMINOSITY = 0.5;

        // color map
        Dictionary<Cluster, Color> assignedColors = new Dictionary<Cluster, Color>();

        /// <summary>
        /// A class that generates colors for clusters.
        /// </summary>
        /// <param name="cs">A clustering.</param>
        /// <param name="degreeStart">The lowest allowable hue, in degrees.</param>
        /// <param name="degreeEnd">The highest allowable hue, in degrees.</param>
        /// <param name="offset">Shift, in degrees.  E.g., if degreeStart = 0 and
        /// degreeEnd = 360 and offset = 45, the effective degreeStart is 45 mod 360 and
        /// the effective degreeEnd is 405 mod 360.</param>
        public ClusterColorer(Clustering cs, double degreeStart, double degreeEnd, double offset)
        {
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
            // also sort clusters so that repainting on subsequent
            // runs produces a stable coloring
            Cluster[] csSorted = cs.OrderByDescending(c => cNeighbors[c].Count).ToArray();

            // greedily assign colors by degree, largest first;
            // aka Welsh-Powell heuristic

            // init angle generator
            var angles = new AngleGenerator(degreeStart, degreeEnd);

            foreach (Cluster c in csSorted)
            {
                // get neighbor colors
                var ns = cNeighbors[c].ToArray();
                var nscs = new HashSet<Color>();
                foreach (Cluster n in ns)
                {
                    if (assignedColors.ContainsKey(n))
                    {
                        nscs.Add(assignedColors[n]);
                    }
                }

                // color getter
                Func<Color> colorf = () =>
                    HSLtoColor(
                        new HSL(
                            mod(
                                angles.NextAngle() + offset,
                                degreeEnd - degreeStart
                            ),
                            SATURATION,
                            LUMINOSITY
                        )
                    );

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

        public static Tuple<Clustering,Dictionary<Cluster,Clustering>> MergeClustersByFingerprint(Clustering cs, ROInvertedHistogram ih)
        {
            var fdict = new Dictionary<Fingerprint, Cluster>();
            var ccdict = new Dictionary<Cluster, Clustering>(); // key is new cluster
            foreach (Cluster c in cs)
            {
                // get location-insensitive fingerprint
                var f = ih[c.First()].Item3.ToCVectorResultant;

                Cluster ckey = null;
                if (fdict.ContainsKey(f))
                {
                    // merge this cluster into existing cluster
                    ckey = fdict[f];
                    ckey.UnionWith(c);
                }
                else
                {
                    // create a new cluster from current cluster
                    ckey = new Cluster();
                    ckey.UnionWith(c);
                    fdict[f] = ckey;
                }

                // update cluster-cluster lookup
                if (!ccdict.ContainsKey(ckey))
                {
                    ccdict[ckey] = new Clustering();
                }
                ccdict[ckey].Add(c);
            }
            var clustering = new HashSet<Cluster>(fdict.Values);

            return new Tuple<Clustering, Dictionary<Cluster, Clustering>>(clustering, ccdict);
        }

        public static Dictionary<Cluster,Clustering> Neighbors(Clustering cs)
        {
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
            return cNeighbors;
        }

        /// <summary>
        /// A class that generates colors for clusters.
        /// </summary>
        /// <param name="clustering">A clustering.</param>
        /// <param name="degreeStart">The lowest allowable hue, in degrees.</param>
        /// <param name="degreeEnd">The highest allowable hue, in degrees.</param>
        /// <param name="offset">Shift, in degrees.  E.g., if degreeStart = 0 and
        /// degreeEnd = 360 and offset = 45, the effective degreeStart is 45 mod 360 and
        /// the effective degreeEnd is 405 mod 360.</param>
        /// <param name="ih">An InvertedHistogram so that coloring is fingerprint-sensitive.</param>
        public ClusterColorer(Clustering clustering, double degreeStart, double degreeEnd, double offset, ROInvertedHistogram ih, Graph g)
        {
            // merge clusters by fingerprint
            var x = MergeClustersByFingerprint(clustering, ih);
            var cs = x.Item1;
            var ccdict = x.Item2;

            // get neighbors
            var neighbors = Neighbors(cs);

            // rank clusters by their degree and break ties using fingerprint (if given fingerprints)
            // also sort clusters so that repainting on subsequent
            // runs produces a stable coloring
            //Cluster[] csSorted =
            //    cs.OrderByDescending(c => neighbors[c].Count)
            //      .ThenBy(c => ih[c.First()].Item3.ToString()).ToArray();
            // size of biggest cluster
            var maxsz = cs.Select(c => c.Count).Max();
            var csSorted = cs.OrderBy(hs =>
            {
                var first = hs.First();
                if (!g.isFormula(first))
                {
                    return 0;
                }
                else
                {
                    return maxsz - hs.Count;
                }
            }).ThenBy(c => ih[c.First()].Item3.ToString()).ToArray();

            // greedily assign colors by degree, largest first;
            // aka Welsh-Powell heuristic

            // init angle generator
            var angles = new AngleGenerator(degreeStart, degreeEnd);

            // color getter
            Func<Color> colorf = () =>
                HSLtoColor(
                    new HSL(
                        mod(
                            angles.NextAngle() + offset,
                            degreeEnd - degreeStart
                        ),
                        SATURATION,
                        LUMINOSITY
                    )
                );

            foreach (Cluster c in csSorted)
            {
                // get neighbor colors
                var ns = neighbors[c].ToArray();
                var nscs = new HashSet<Color>();
                foreach (Cluster n in ns)
                {
                    if (assignedColors.ContainsKey(n))
                    {
                        nscs.Add(assignedColors[n]);
                    }
                }

                // get initial color
                var color = colorf();
                while (nscs.Contains(color))
                {
                    // if we've already chosen this one, get next color
                    color = colorf();
                }

                // save this color for each _original_ cluster in c
                foreach(Cluster corig in ccdict[c])
                {
                    assignedColors[corig] = color;
                }
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
            var hs = c.SelectMany(addr => ExceLint.Utils.AdjacentCells(addr))
                      .Aggregate(new HashSet<AST.Address>(), (acc, addr) => {
                          acc.Add(addr);
                          return acc;
                       });
            return hs;
        }

        public Color GetColor(Cluster c)
        {
            return assignedColors[c];
        }
    }
}
