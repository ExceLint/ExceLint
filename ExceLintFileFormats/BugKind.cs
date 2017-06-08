using System;

namespace ExceLintFileFormats
{
    public class UnknownBugType : Exception
    {
        public UnknownBugType(string message) : base(message) { }
    }

    public abstract class BugKind
    {
        public abstract string ToLog();

        public static BugKind ToKind(string kindstr)
        {
            switch (kindstr)
            {
                case "no":
                    return NotABug;
                case "ref":
                    return ReferenceBug;
                case "refi":
                    return ReferenceBugInverse;
                case "calc":
                    return CalculationError;
                case "cwfe":
                    return ConstantWhereFormulaExpected;
                case "fwce":
                    return FormulaWhereConstantExpected;
                case "pmi":
                    return PotentialMaintenanceIssue;
                case "iuf":
                    return InconsistentUseOfFormula;
                default:
                    throw new UnknownBugType(kindstr);
            }
        }

        public static BugKind[] AllKinds
        {
            get
            {
                return new BugKind[] {
                    NotABug,
                    ReferenceBug,
                    ReferenceBugInverse,
                    CalculationError,
                    ConstantWhereFormulaExpected,
                    FormulaWhereConstantExpected,
                    PotentialMaintenanceIssue,
                    InconsistentUseOfFormula
                };
            }
        }

        public static NotABug NotABug
        {
            get { return NotABug.Instance; }
        }

        public static ReferenceBug ReferenceBug
        {
            get { return ReferenceBug.Instance; }
        }

        public static ReferenceBugInverse ReferenceBugInverse
        {
            get { return ReferenceBugInverse.Instance; }
        }

        public static CalculationError CalculationError
        {
            get { return CalculationError.Instance; }
        }

        public static ConstantWhereFormulaExpected ConstantWhereFormulaExpected
        {
            get { return ConstantWhereFormulaExpected.Instance; }
        }

        public static FormulaWhereConstantExpected FormulaWhereConstantExpected
        {
            get { return FormulaWhereConstantExpected.Instance; }
        }

        public static PotentialMaintenanceIssue PotentialMaintenanceIssue
        {
            get { return PotentialMaintenanceIssue.Instance; }
        }

        public static InconsistentUseOfFormula InconsistentUseOfFormula
        {
            get { return InconsistentUseOfFormula.Instance; }
        }

        public static BugKind DefaultKind
        {
            get { return NotABug; }
        }
    }

    public class NotABug : BugKind
    {
        private static NotABug instance;

        private NotABug() { }

        public static NotABug Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new NotABug();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "no";
        }

        public override string ToString()
        {
            return "Not a bug";
        }
    }

    public class ReferenceBug : BugKind
    {
        private static ReferenceBug instance;

        private ReferenceBug() { }

        public static ReferenceBug Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ReferenceBug();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "ref";
        }

        public override string ToString()
        {
            return "Reference bug";
        }
    }

    public class ReferenceBugInverse : BugKind
    {
        private static ReferenceBugInverse instance;

        private ReferenceBugInverse() { }

        public static ReferenceBugInverse Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ReferenceBugInverse();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "refi";
        }

        public override string ToString()
        {
            return "Reference bug (inverse)";
        }
    }

    public class CalculationError : BugKind
    {
        private static CalculationError instance;

        private CalculationError() { }

        public static CalculationError Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new CalculationError();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "calc";
        }

        public override string ToString()
        {
            return "Calculation error";
        }
    }

    public class ConstantWhereFormulaExpected : BugKind
    {
        private static ConstantWhereFormulaExpected instance;

        private ConstantWhereFormulaExpected() { }

        public static ConstantWhereFormulaExpected Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ConstantWhereFormulaExpected();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "cwfe";
        }

        public override string ToString()
        {
            return "Constant where formula expected";
        }
    }

    public class FormulaWhereConstantExpected : BugKind
    {
        private static FormulaWhereConstantExpected instance;

        private FormulaWhereConstantExpected() { }

        public static FormulaWhereConstantExpected Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new FormulaWhereConstantExpected();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "fwce";
        }

        public override string ToString()
        {
            return "Formula where constant expected";
        }
    }

    public class PotentialMaintenanceIssue : BugKind
    {
        private static PotentialMaintenanceIssue instance;

        private PotentialMaintenanceIssue() { }

        public static PotentialMaintenanceIssue Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new PotentialMaintenanceIssue();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "pmi";
        }

        public override string ToString()
        {
            return "Potential maintenance issue";
        }
    }

    public class InconsistentUseOfFormula : BugKind
    {
        private static InconsistentUseOfFormula instance;

        private InconsistentUseOfFormula() { }

        public static InconsistentUseOfFormula Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new InconsistentUseOfFormula();
                }
                return instance;
            }
        }

        public override string ToLog()
        {
            return "iuf";
        }

        public override string ToString()
        {
            return "Inconsistent use of formula";
        }
    }
}
