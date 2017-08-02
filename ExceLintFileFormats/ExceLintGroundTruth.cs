using System.IO;
using System.Collections.Generic;
using System.Linq;
using CsvHelper;
using System;

namespace ExceLintFileFormats
{
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

    public class ExceLintGroundTruth: IDisposable
    {
        private string _dbpath;
        private Dictionary<AST.Address, BugKind> _bugs = new Dictionary<AST.Address, BugKind>();
        private Dictionary<AST.Address, string> _notes = new Dictionary<AST.Address, string>();
        private HashSet<AST.Address> _added = new HashSet<AST.Address>();
        private HashSet<AST.Address> _changed = new HashSet<AST.Address>();

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

        private ExceLintGroundTruth(string dbpath, ExceLintGroundTruthRow[] rows)
        {
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

        public List<System.Tuple<AST.Address,BugAnnotation>> AnnotationsFor(string workbookname)
        {
            var output = new List<System.Tuple<AST.Address, BugAnnotation>>();

            foreach (var bug in _bugs)
            {
                var addr = bug.Key;

                if (addr.WorkbookName == workbookname)
                {
                    output.Add(new System.Tuple<AST.Address, BugAnnotation>(addr, new BugAnnotation(bug.Value, _notes[bug.Key])));
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

        public HashSet<AST.Address> TrueRefBugsByWorkbook(string wbname)
        {
            return new HashSet<AST.Address>(
                _bugs
                    .Where(pair => pair.Key.A1Workbook() == wbname)
                    .Select(pair => pair.Key)
                );
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
