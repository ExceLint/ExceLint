using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExceLintUI
{
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
    }
}
