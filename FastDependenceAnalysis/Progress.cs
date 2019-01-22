using System;

namespace FastDependenceAnalysis
{
    public delegate void ProgressBarIncrementer(int n);
    public delegate void ProgressBarReset();

    public class Progress
    {
        private volatile bool _cancelled = false;
        private long _total = 0;
        private ProgressBarIncrementer _progBarIncr;
        private ProgressBarReset _progBarReset;
        private long _workMultiplier = 1;
        private long _current = 0;

        public static Progress NOPProgress()
        {
            ProgressBarIncrementer pbi = (int n) => { return; };
            ProgressBarReset pbr = () => { return; };
            return new Progress(pbi, pbr, 1L);
        }

        public Progress(ProgressBarIncrementer progBarIncrement, ProgressBarReset progBarReset, long workMultiplier)
        {
            _progBarIncr = progBarIncrement;
            _progBarReset = progBarReset;
            _workMultiplier = workMultiplier;
        }

        public long TotalWorkUnits
        {
            get { return _total; }
            set { _total = value; }
        }

        public long UpdateEvery
        {
            get { return Math.Max(1L, (_total * _workMultiplier) / 100L); }
        }

        public void IncrementCounter()
        {
            IncrementCounterN(1);
        }

        public void IncrementCounterN(int n)
        {
            _current += n;
            if (_current % UpdateEvery == 0)
            {
                _progBarIncr(n);
            }
        }

        public void Cancel()
        {
            _cancelled = true;
        }

        public bool IsCancelled()
        {
            return _cancelled;
        }

        public void Reset()
        {
            _progBarReset();
        }
    }
}
