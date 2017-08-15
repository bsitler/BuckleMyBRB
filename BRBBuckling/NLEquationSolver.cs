using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BRBBuckling
{
    internal static class NLEquationSolver
    {
        internal static double BruteForce(Func<double,double> eqn, double target, double minVal, double maxVal, double step, int root = 1)
        {

            double val = eqn(minVal);
            if (val == double.NaN) return double.NaN;
            else if (val == eqn(maxVal)) return double.NaN;

            double sign = val- target>0 ? 1 : -1;

            for (double N = minVal; N < maxVal; N += step)
            {
                val = eqn(N);
                if (val== double.NaN)
                {
                    return double.NaN;
                }
                else if ((val-target)* sign < 0)
                {
                    root--; sign = -sign;
                    if (root < 1) return N - step;
                }
            }
            return double.NaN;
        }
    }
}
