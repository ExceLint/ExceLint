using System.IO;
using System.Collections.Generic;
using System.Linq;
using CsvHelper;
using System;
using System.Text.RegularExpressions;
using BugClass = System.Collections.Generic.HashSet<AST.Address>;
using Microsoft.FSharp.Core;

namespace ExceLintFileFormats
{
    public class AddressComparer : IComparer<AST.Address>
    {
        public int Compare(AST.Address a1, AST.Address a2)
        {
            if (a1.Y < a2.Y)
            {
                return -1;
            }
            else if (a1.Y == a2.Y && a1.X < a2.X)
            {
                return -1;
            }
            else if (a1.Y == a2.Y && a1.X == a2.X)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }
    }

    public struct BugAnnotation
    {
        public BugKind BugKind;
        public string Note;

        public BugAnnotation(BugKind bugKind, string note)
        {
            BugKind = bugKind;
            Note = note;
        }

        public string Comment
        {
            get { return BugKind.ToString() + ": " + Note; }
        }
    };

    public class ExceLintGroundTruth : IDisposable
    {
        private string _dbpath;
        private Dictionary<AST.Address, BugKind> _bugs = new Dictionary<AST.Address, BugKind>();
        private Dictionary<AST.Address, string> _notes = new Dictionary<AST.Address, string>();
        private HashSet<AST.Address> _added = new HashSet<AST.Address>();
        private HashSet<AST.Address> _changed = new HashSet<AST.Address>();
        private Dictionary<AST.Address, BugClass> _bugclass_lookup = new Dictionary<AST.Address, BugClass>();
        private Dictionary<AST.Address, BugClass> _dual_lookup = new Dictionary<AST.Address, BugClass>();
        private Dictionary<BugClass, BugClass> _bugclass_dual_lookup = new Dictionary<BugClass, BugClass>();
        private HashSet<string> _has_annotations_for_workbook = new HashSet<string>();

        /// <summary>
        /// Get the AST.Address for a row.
        /// </summary>
        /// <param name="addrStr"></param>
        /// <param name="worksheetName"></param>
        /// <param name="workbookName"></param>
        /// <returns></returns>
        private AST.Address Address(string addrStr, string worksheetName, string workbookName)
        {
            return AST.Address.FromA1String(
                addrStr.ToUpper(),
                worksheetName,
                workbookName,
                ""  // we don't care about paths
            );
        }

        private FSharpOption<BugClass> DualsFor(AST.Address addr)
        {
            // extract address environment
            var env = new AST.Env(addr.Path, addr.WorkbookName, addr.WorksheetName);

            // duals regexp
            var r = new Regex(@".*dual\s*=\s*((?<AddrOrRange>[A-Z]+[0-9]+(:?:[A-Z]+[0-9]+)?)(:?\s*,\s*)?)+", RegexOptions.Compiled);

            // get note for this address
            var note = _notes[addr];

            Match m = r.Match(note);
            if (!m.Success)
            {
                if (note.Contains("dual"))
                {
                    Console.Out.WriteLine("Malformed dual annotation for cell " + addr.A1FullyQualified() + " : " + note);
                }
                return FSharpOption<BugClass>.None;
            } else
            {
                // init duals list
                var duals = new List<AST.Address>();

                var cs = m.Groups["AddrOrRange"].Captures;

                foreach (Capture c in cs)
                {
                    // get string value
                    string addrOrRange = c.Value;

                    AST.Reference xlref = null;

                    try
                    {
                        // parse
                        xlref = Parcel.simpleReferenceParser(addrOrRange, env);
                    } catch (Exception e)
                    {
                        var msg = "Bad reference: '" + addrOrRange + "'";
                        Console.Out.WriteLine(msg);
                        throw new Exception(msg);
                    }

                    // figure out the reference type
                    if (xlref.Type == AST.ReferenceType.ReferenceRange)
                    {
                        var rrref = (AST.ReferenceRange)xlref;
                        duals.AddRange(rrref.Range.Addresses());
                    } else if (xlref.Type == AST.ReferenceType.ReferenceAddress)
                    {
                        var aref = (AST.ReferenceAddress)xlref;
                        duals.Add(aref.Address);
                    } else
                    {
                        throw new Exception("Unsupported address reference type.");
                    }
                }

                var bugclass = new BugClass(duals);

                return FSharpOption<BugClass>.Some(bugclass);
            }
        }

        private ExceLintGroundTruth(string dbpath, ExceLintGroundTruthRow[] rows)
        {
            Console.WriteLine("Indexing ExceLint bug database...");

            _dbpath = dbpath;

            foreach (var row in rows)
            {
                if (row.Address != "Address")
                {
                    AST.Address addr = Address(row.Address, row.Worksheet, row.Workbook);
                    _bugs.Add(addr, BugKind.ToKind(row.BugKind));
                    _notes.Add(addr, row.Notes);
                }
            }

            // index workbooks
            _has_annotations_for_workbook = new HashSet<string>(_bugs.Keys.Select(a => a.WorkbookName).Distinct());

            // find all bugclasses
            foreach (KeyValuePair<AST.Address, BugKind> kvp in _bugs)
            {
                // get address
                var addr = kvp.Key;

                // get duals, if there are any
                var dual_opt = DualsFor(addr);
                if (FSharpOption<BugClass>.get_IsSome(dual_opt))
                {
                    // dual bugclass
                    var duals = dual_opt.Value;

                    // store each dual address in a bugclass if it hasn't already been stored
                    foreach (AST.Address caddr in duals)
                    {
                        // if no bugclass is stored for address
                        if (!_bugclass_lookup.ContainsKey(caddr))
                        {
                            // add it
                            _bugclass_lookup.Add(caddr, duals);
                        }
                    }

                    // get all the addresses in dual and saved bugclasses
                    var classaddrs = duals.SelectMany(caddr => _bugclass_lookup[caddr]).Distinct().ToArray();

                    // get an arbitrary bugclass
                    var fstbugclass = _bugclass_lookup[classaddrs.First()];

                    // is every address in this class?  if not, add them
                    foreach (AST.Address caddr in classaddrs)
                    {
                        if (!fstbugclass.Contains(caddr))
                        {
                            fstbugclass.Add(caddr);
                        }
                    }

                    // ensure that every address refers to the very same bugclass object
                    foreach (AST.Address caddr in classaddrs)
                    {
                        _bugclass_lookup[caddr] = fstbugclass;
                    }

                    // now make sure that the dual lookup for addr points to fstbugclass
                    _dual_lookup.Add(addr, fstbugclass);
                }
            }

            // make sure that all of the bugs in the dual bugclasses have lookups for their own bugclasses
            foreach (var kvp in _dual_lookup)
            {
                var bugclass = kvp.Value;

                foreach (var addr in bugclass)
                {
                    if (!_bugclass_lookup.ContainsKey(addr))
                    {
                        _bugclass_lookup.Add(addr, bugclass);
                    }
                }
            }

            // make sure that every bug remaining (i.e., those without duals) is in a bugclass
            foreach (var kvp in _bugs)
            {
                var addr = kvp.Key;
                if (!_bugclass_lookup.ContainsKey(addr))
                {
                    var bc = new HashSet<AST.Address>();
                    bc.Add(addr);
                    _bugclass_lookup.Add(addr, bc);
                }
            }

            // now index bugclass -> bugclass dual lookup
            foreach (var kvp in _bugclass_lookup)
            {
                var addr = kvp.Key;
                var bugclass = kvp.Value;

                // did we already save the dual for this bugclass?
                if (!_bugclass_dual_lookup.ContainsKey(bugclass))
                {
                    // grab the dual bugclass for this bugclass, if it has one
                    if (_dual_lookup.ContainsKey(addr))
                    {
                        var dual = _dual_lookup[addr];

                        // since the bugclass should be the same for all
                        // addresses in the bugclass, just lookup the bugclass
                        // by an arbitrary representative address
                        var fstdual = dual.First();
                        var dual_bugclass = _bugclass_lookup[fstdual];

                        // now save it
                        _bugclass_dual_lookup.Add(bugclass, dual_bugclass);
                    }
                }
            }

            // sanity checks

            // every bug mentioned in notes has an entry in the lookup table
            foreach (var kvp in _bugclass_lookup)
            {
                var addr = kvp.Key;
                if (!_bugs.ContainsKey(addr))
                {
                    Console.WriteLine("WARNING: Address " + addr.A1FullyQualified() + " referenced in bug notes but not annotated.");
                }
            }

            // all duals are mutually exclusive
            foreach (var kvp in _bugclass_dual_lookup)
            {
                BugClass bc1 = kvp.Key;
                BugClass bc2 = kvp.Value;

                if (bc1.Intersect(bc2).Count() > 0)
                {
                    string bc1str = String.Join(",", bc1.Select(a => a.A1Local()));
                    string bc2str = String.Join(",", bc2.Select(a => a.A1Local()));
                    string wb = bc1.First().A1Workbook();
                    string ws = bc1.First().A1Worksheet();
                    Console.WriteLine(
                        "WARNING: bug class\n\t" +
                        bc1str +
                        "\n\tis not mutually exclusive with bug class\n\t" +
                        bc2str +
                        "\n\tin workbook " + wb + " on worksheet " + ws);
                }
            }

            Console.WriteLine("Done indexing ExceLint bug database.");
        }

        public BugAnnotation AnnotationFor(AST.Address addr)
        {
            if (_bugs.ContainsKey(addr))
            {
                return new BugAnnotation(_bugs[addr], _notes[addr]);
            } else
            {
                return new BugAnnotation(BugKind.NotABug, "");
            }
        }

        public List<Tuple<AST.Address, BugAnnotation>> AnnotationsFor(string workbookname)
        {
            var output = new List<Tuple<AST.Address, BugAnnotation>>();

            foreach (var bug in _bugs)
            {
                var addr = bug.Key;

                if (addr.WorkbookName == workbookname)
                {
                    output.Add(new Tuple<AST.Address, BugAnnotation>(addr, new BugAnnotation(bug.Value, _notes[bug.Key])));
                }
            }

            return output;
        }

        /// <summary>
        /// Insert or update an annotation for a given cell.
        /// </summary>
        /// <param name="addr"></param>
        /// <param name="annot"></param>
        public void SetAnnotationFor(AST.Address addr, BugAnnotation annot)
        {
            if (_bugs.ContainsKey(addr))
            {
                _bugs[addr] = annot.BugKind;
                _notes[addr] = annot.Note;
                _changed.Add(addr);
            }
            else
            {
                _bugs.Add(addr, annot.BugKind);
                _notes.Add(addr, annot.Note);
                _added.Add(addr);
            }
        }

        public void Write()
        {
            // always append if any existing line changed
            var noAppend = _changed.Count() > 0;

            using (StreamWriter sw = new StreamWriter(
                path: _dbpath,
                append: !noAppend,
                encoding: System.Text.Encoding.UTF8))
            {
                using (CsvWriter cw = new CsvWriter(sw))
                {
                    // if any annotation changed, we need to write
                    // the entire file over again.
                    if (noAppend)
                    {
                        // write header
                        cw.WriteHeader<ExceLintGroundTruthRow>();

                        // write the rest of the file
                        foreach (var pair in _bugs)
                        {
                            var addr = pair.Key;
                            var bugAnnotation = AnnotationFor(addr);

                            var row = new ExceLintGroundTruthRow();
                            row.Address = addr.A1Local();
                            row.Worksheet = addr.A1Worksheet();
                            row.Workbook = addr.A1Workbook();
                            row.BugKind = bugAnnotation.BugKind.ToLog();
                            row.Notes = bugAnnotation.Note;

                            cw.WriteRecord(row);
                        }

                        _changed.Clear();
                        _added.Clear();
                    } else
                    {
                        // if the total number of bugs is not the same
                        // as the number of changes, then it's because 
                        // some annotations came from a file and thus
                        // we are only appending; do not write a new
                        // header when appending because we already 
                        // have one.
                        if (_bugs.Count() == _added.Count())
                        {
                            cw.WriteHeader<ExceLintGroundTruthRow>();
                        }

                        foreach (var addr in _added)
                        {
                            var bugAnnotation = AnnotationFor(addr);

                            var row = new ExceLintGroundTruthRow();
                            row.Address = addr.A1Local();
                            row.Worksheet = addr.A1Worksheet();
                            row.Workbook = addr.A1Workbook();
                            row.BugKind = bugAnnotation.BugKind.ToLog();
                            row.Notes = bugAnnotation.Note;

                            cw.WriteRecord(row);
                        }

                        _added.Clear();
                    }
                }
            }
        }

        public bool IsABug(AST.Address addr)
        {
            return _bugs.ContainsKey(addr) && _bugs[addr] != BugKind.NotABug;
        }

        public bool IsATrueRefBug(AST.Address addr)
        {
            return _bugs.ContainsKey(addr) && IsTrueRefBug(_bugs[addr]);
        }

        public bool IsATrueRefBugOrSuspicious(AST.Address addr)
        {
            return _bugs.ContainsKey(addr) && IsTrueRefBugOrSuspicious(_bugs[addr]);
        }

        private bool IsTrueRefBug(BugKind b)
        {
            return
                b == BugKind.FormulaWhereConstantExpected ||
                b == BugKind.ConstantWhereFormulaExpected ||
                b == BugKind.ReferenceBug ||
                b == BugKind.ReferenceBugInverse ||
                b == BugKind.CalculationError;
        }

        private bool IsTrueRefBugOrSuspicious(BugKind b)
        {
            return
                IsTrueRefBug(b) ||
                b == BugKind.SuspiciousCell;
        }

        private AST.Address LeftTopForBugClass(BugClass bc)
        {
            var ordered = bc.OrderBy(a => a, new AddressComparer());
            return ordered.First();
        }

        public bool AddressHasADual(AST.Address addr)
        {
            if (_bugclass_lookup.ContainsKey(addr))
            {
                var bugclass = _bugclass_lookup[addr];
                return _bugclass_dual_lookup.ContainsKey(bugclass);
            } else
            {
                return false;
            }
        }

        public Tuple<BugClass,BugClass> DualsForAddress(AST.Address addr)
        {
            var bugclass = _bugclass_lookup[addr];
            var bugclass_dual = _bugclass_dual_lookup[bugclass];

            // which bugclass comes first? order by class with the topleftmost topleft corner
            var bc_lt = LeftTopForBugClass(bugclass);
            var bd_lt = LeftTopForBugClass(bugclass_dual);

            var cmp = new AddressComparer();

            var tup = cmp.Compare(bc_lt, bd_lt) < 0 ? 
                      new Tuple<BugClass, BugClass>(bugclass, bugclass_dual) :
                      new Tuple<BugClass, BugClass>(bugclass_dual, bugclass);

            return new Tuple<BugClass,BugClass>(bugclass,_bugclass_dual_lookup[bugclass]);
        }

        public bool HasTrueRefAnnotations(string wbname)
        {
            return _has_annotations_for_workbook.Contains(wbname);
        }

        public string[] WorkbooksAnnotated
        {
            get { return _bugs.Keys.Select(a => a.WorkbookName).Distinct().ToArray(); }
        }

        public int TotalNumTrueRefBugs
        {
            get
            {
                var wbs = WorkbooksAnnotated;
                int count = 0;
                foreach (var wb in wbs)
                {
                    count += NumTrueRefBugsForWorkbook(wb);
                }
                return count;
            }
        }

        public int NumTrueRefBugsForWorkbook(string wbname)
        {
            // get the set of duals relevant for this workbook
            var duals =
                _bugclass_dual_lookup
                .Where(kvp =>
                    kvp.Key.First().WorkbookName == wbname &&   // where the workbook name is the same
                    IsTrueRefBug(_bugs[kvp.Key.First()]));      // and it's actually a reference bug

            // eliminate converse duals
            var duals_nodupes = new Dictionary<BugClass,BugClass>();
            foreach (var kvp in duals)
            {
                var bc = kvp.Key;
                var dualbc = kvp.Value;
                if (!duals_nodupes.ContainsKey(bc) && !duals_nodupes.ContainsKey(dualbc))
                {
                    duals_nodupes.Add(bc, dualbc);
                }
            }

            // flatten
            var duals_addrs = new HashSet<AST.Address>();
            foreach (var kvp in duals_nodupes)
            {
                foreach (var addr in kvp.Key)
                {
                    duals_addrs.Add(addr);
                }
                foreach (var addr in kvp.Value)
                {
                    duals_addrs.Add(addr);
                }
            }

            // now get all bugs that don't have duals
            var nodual_bugs = _bugs.Where(kvp =>
                kvp.Key.WorkbookName == wbname &&       // where the workbook name matches
                !duals_addrs.Contains(kvp.Key));        // and the bug has no dual

            // now count for duals
            int bugs = 0;
            foreach (var kvp in duals_nodupes)
            {
                bugs += NumBugsForBugClass(kvp.Key);
            }

            // and count non-dual bugs
            bugs += nodual_bugs.Count();

            return bugs;
        }

        public int NumBugsForBugClass(BugClass bc)
        {
            var dualbc = _bugclass_dual_lookup[bc];
            return Math.Min(bc.Count, dualbc.Count);
        }

        public HashSet<AST.Address> Flags
        {
            get { return new HashSet<AST.Address>(_bugs.Keys); }
        }

        public static ExceLintGroundTruth Load(string path)
        {
            using (var sr = new StreamReader(path))
            {
                var rows = new CsvReader(sr).GetRecords<ExceLintGroundTruthRow>().ToArray();

                return new ExceLintGroundTruth(path, rows);
            }
        }

        public static ExceLintGroundTruth Create(string gtpath)
        {
            using (StreamWriter sw = new StreamWriter(gtpath))
            {
                using (CsvWriter cw = new CsvWriter(sw))
                {
                    cw.WriteHeader<ExceLintGroundTruthRow>();
                }
            }

            return Load(gtpath);
        }

        public void Dispose()
        {
            Write();
        }
    }

    class ExceLintGroundTruthRow
    {
        public string Workbook { get; set; }
        public string Worksheet { get; set; }
        public string Address { get; set; }
        public string BugKind { get; set; }
        public string Notes { get; set; }
    }
}
