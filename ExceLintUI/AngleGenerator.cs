using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;

namespace ExceLintUI
{
    public struct Arc
    {
        public double start;
        public double end;

        public Arc(double start, double end)
        {
            this.start = start;
            this.end = end;
        }

        public double Start
        {
            get { return start; }
        }
        public double End
        { 
            get { return end; }
        }
    }

    public class AngleGenerator
    {
        IEnumerable<double> angles;

        public AngleGenerator(double start, double end)
        {
            angles = Angles(start, end);
        }

        public double NextAngle()
        {
            var a = angles.Take(1).First();
            angles = angles.Skip(1);
            return a;
        }

        private static IEnumerable<double> OldAngles(double start, double end)
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

        private static IEnumerable<double> Angles(double start, double end)
        {
            Func<double, double, double> midpoint_f = (s, e) => (e - s) / 2 + s;

            var work = new Stack<Arc>();

            // initialize
            work.Push(new Arc(start, end));

            while (true)
            {
                // grab job, compute midpoint and yield
                var job = work.Pop();
                var midpoint = midpoint_f(job.start, job.end);
                yield return midpoint;

                // put next two arcs on stack
                work.Push(new Arc(midpoint, end));
                work.Push(new Arc(start, midpoint));
            }
        }
    }
}
