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
                case "cwfe":
                    return ConstantWhereFormulaExpected;
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
                    ConstantWhereFormulaExpected
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

        public static ConstantWhereFormulaExpected ConstantWhereFormulaExpected
        {
            get { return ConstantWhereFormulaExpected.Instance; }
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
}
