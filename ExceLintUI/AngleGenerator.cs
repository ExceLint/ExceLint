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

        private static IEnumerable<double> Angles(double start, double end)
        {
            Func<double, double, double> midpoint_f = (s, e) => (e - s) / 2 + s;
            var work = new Queue<Arc>();

            // initialize
            work.Enqueue(new Arc(start, end));

            while (true)
            {
                // grab job, compute midpoint and yield
                var job = work.Dequeue();
                var midpoint = midpoint_f(job.start, job.end);
                yield return midpoint;

                // put next two arcs on queue
                work.Enqueue(new Arc(start, midpoint));
                work.Enqueue(new Arc(midpoint, end));
            }
        }
    }
}
