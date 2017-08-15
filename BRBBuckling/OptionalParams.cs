using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BRBBuckling
{
    internal static class OptionalParams
    {
        internal static T Check<T>(object arg, T defaultValue)
        {
            if (arg is T)
                return (T)arg;
            else return defaultValue;
        }
        internal static double ToDoubleInfinity(object val)
        {
            // double inf = 8734;
            // return (AscW(val) == inf);
            if (val.ToString() == "∞") return double.PositiveInfinity;
            else if (val is double) return (double)val;
            else return double.NaN;
        }
    }
}
