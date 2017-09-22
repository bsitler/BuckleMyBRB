using System;
using static System.Math;
using linalg = MathNet.Numerics.LinearAlgebra;

internal static class ElasticBuckling
{
    #region Elastic buckling load
    /// <summary>ke=1/√(NcrB/Ncr0), or NcrB = π²γEIr/(ke·ξL0)²=Ncr0/ke²</summary>
    /// <param name="dN">increment of N/Ncr0</param>
    /// <param name="Nmax">max N/Ncr0</param>
    internal static double ke_crit(Func<double, double> det_func, double modeNum = 1, double dN = 0.0001, double Nmax = 1.0)
    {
        //If IsError(det)Then
        //    NcreSym = det
        //    Eξt Function
        //End If

        double det = det_func(dN);
        if (det == double.NaN) return double.NaN;
        else if (det == 0) return double.PositiveInfinity;

        double sign = det > 0 ? 1 : -1;

        for (double N = dN; N < Nmax; N += dN)
        {
            det = det_func(N);
            if (det == double.NaN)
            {
                return double.NaN;
            }
            else if (det * sign < 0)
            {
                modeNum--; sign = -sign;
                if (modeNum < 1)
                    return 1 / Sqrt(N - dN);
            }
        }
        return 0;
    }
    internal static double det_sym(double N, double κg, double κr, double ξ, double γ)
    {
        if (N < 0 || γ <= 0 || ξ <= 0 || (κg <= 0 && κr <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double S2 = Sin(Sqrt(N) * PI * Sqrt(γ) / ξ * (0.5 - ξ));
        double C2 = Cos(Sqrt(N) * PI * Sqrt(γ) / ξ * (0.5 - ξ));

        return N * PI * PI * S1 * C2 + Sqrt(N) * PI *
          (κr * (Sqrt(γ) * S1 * S2 - C1 * C2) - κg * C1 * C2)
            - κg * κr * (S1 * C2 + Sqrt(γ) * C1 * S2);
    }
    internal static double det_asym(double N, double κg, double κr, double ξ, double γ)
    {
        if (N < 0 || γ <= 0 || ξ <= 0 || (κg <= 0 && κr <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double S2 = Sin(Sqrt(N) * PI * Sqrt(γ) / ξ * (0.5 - ξ));
        double C2 = Cos(Sqrt(N) * PI * Sqrt(γ) / ξ * (0.5 - ξ));

        return Sqrt(N) * N * PI * PI * PI * S1 * S2
        - N * PI * PI * (κr * (Sqrt(γ) * S1 * C2 + C1 * S2) + κg * C1 * S2)
        + PI * Sqrt(N) * κg * (2 * ξ * S1 * S2 + κr * (Sqrt(γ) * C1 * C2 - S1 * S2))
        - 2 * κr * κg * ξ * (Sqrt(γ) * S1 * C2 + C1 * S2);
    }
    internal static double det_one(double N, double κg, double κr, double ξ, double γ)
    {
        if (N < 0 || γ <= 0 || ξ <= 0 || (κg <= 0 && κr <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double S10 = Sin(Sqrt(N) * PI * Sqrt(γ) * (1 - 1 / ξ));
        double C10 = Cos(Sqrt(N) * PI * Sqrt(γ) * (1 - 1 / ξ));

        return Sqrt(N) * N * PI * PI * PI * S1 * S10
        + PI * PI * N * (κr * (Sqrt(γ) * S1 * C10 - C1 * S10) - κg * C1 * S10)
        + PI * κg * Sqrt(N) * (ξ * S1 * S10 - κr * (Sqrt(γ) * C1 * C10 + S1 * S10))
        + κr * κg * ξ * (Sqrt(γ) * S1 * C10 - C1 * S10);
    }
    internal static double det_chev(double N, double κg1, double κg2, double κr1, double κr2,
         double ξ1, double ξ2, double γ1, double γ2)
    {
        if (N < 0 || γ1 <= 0 || γ2 <= 0 || ξ1 <= 0 || ξ2 <= 0 
            || (κg1 <= 0 && κr1 <= 0) || (κg2 <= 0 && κr2 <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double S4 = Sin(Sqrt(N) * Sqrt(γ1) * PI);
        double C4 = Cos(Sqrt(N) * Sqrt(γ1) * PI);
        double S6 = Sin(Sqrt(N) * Sqrt(γ1 / γ2) * PI / ξ1 * (1 - ξ2));
        double C6 = Cos(Sqrt(N) * Sqrt(γ1 / γ2) * PI / ξ1 * (1 - ξ2));
        double S7 = Sin(Sqrt(N) * Sqrt(γ1) * PI / ξ1 * (1 - ξ2));
        double C7 = Cos(Sqrt(N) * Sqrt(γ1) * PI / ξ1 * (1 - ξ2));
        double S8 = Sin(Sqrt(N) * Sqrt(γ1 / γ2) * PI / ξ1);
        double C8 = Cos(Sqrt(N) * Sqrt(γ1 / γ2) * PI / ξ1);

        var M = linalg.Double.DenseMatrix.OfArray(new double[,] {
          { κg1 * PI / ξ1 * Sqrt(N), κg1 + N * Pow(PI,2) / ξ1, 0, 0, -κg1 * S8, -κg1 * C8 },
          { S1, C1, -S4, -C4, 0, 0},
          { -κr1 * C1, κr1* S1, Sqrt(N) *PI * S4 + Sqrt(γ1) * κr1 * C4, Sqrt(N)* PI *C4 - Sqrt(γ1) * κr1 * S4, 0, 0},
          { 0, 0, S7, C7, -S6, -C6},
          { 0, 0, -Sqrt(γ2) * κr2 * C7, Sqrt(γ2) * κr2 * S7, Sqrt(N)* PI *ξ2 / ξ1 * Sqrt(γ1 / γ2) * S6 + κr2 * C6,
                     Sqrt(N)* PI *ξ2 / ξ1 * Sqrt(γ1 / γ2) * C6 - κr2 * S6},
          { Sqrt(γ2 / γ1) * κg1*κg2, Sqrt(N)* Sqrt(γ2 / γ1) * PI * κg2, 0, 0,
                     κg1*Sqrt(N) * PI * ξ2 / ξ1 * Sqrt(γ1 / γ2) * S8 - κg1*κg2 * C8,
                     κg1*Sqrt(N) * PI * ξ2 / ξ1 * Sqrt(γ1 / γ2) * C8 + κg1*κg2 * S8} });

        if (κg1==0 && κg2==0) M.SetRow(5, new double[] {
            0, γ2 / γ1 , 0, 0,ξ2 / ξ1 * S8,ξ2 / ξ1 * C8});

        return M.Determinant();
    }


    internal static double det_sym_rigidrest(double N, double κg, double κr, double ξ, double γ)
    {
        if (N < 0 || γ <= 0 || ξ <= 0 || (κg <= 0 && κr <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double n = PI * Sqrt(N);

        // combined spring including restrainer stiffness
        κr = 1.0 / (1 / κr + γ*(1 - 2 * ξ) / 2/ ξ);

        return κg * (n * C1 + κr * S1) - n * (n * S1 - κr * C1);
    }
    internal static double det_asym_rigidrest(double N, double κg, double κr, double ξ, double γ)
    {
        if (N < 0 || γ <= 0 || ξ <= 0 || (κg <= 0 && κr <= 0)) return 0;
        double S1 = Sin(Sqrt(N) * PI);
        double C1 = Cos(Sqrt(N) * PI);
        double n = PI * Sqrt(N);

        // combined spring including restrainer stiffness
        κr =  1.0 / (1 / κr + γ*(1 - 2 * ξ) / 6/ ξ);
        //double κtr = 24 * Pow(ξ, 3) / γ / Pow(1 - 2 * ξ, 3);

        //var M = linalg.Double.DenseMatrix.OfArray(new double[,] {
        //    { κg,            n,            κg/n            },
        //    { κtr*S1,        κtr*(C1-1),    κtr-Pow(n,2)    },
        //    { n*S1-κr*C1,    n*C1+κr*S1,    -κr/n            }});
        //var M = linalg.Double.DenseMatrix.OfArray(new double[,] {
        //    { κg,           n,          κg/n              },
        //    { n*S1-2*κr*C1,   n*C1+2*κr*S1, 0               },
        //    { 2*ξ*(S1+n*C1-1)-n*C1, 2*ξ*(C1-n*S1)+n*S1,0   }});
        var M = linalg.Double.DenseMatrix.OfArray(new double[,] {
            { 0,1,0,1,0,0 },
            { κg,n,κg/n,0,0,0},
            {S1,C1,1,1,-1,-1 },
            {0,0,1,0,-1,0},
            {n*S1-κr*C1,n*C1+κr*S1,-κr/n,0,κr/n,0},
            {0,0,0,0,1.0/2/ξ,1}});

        return M.Determinant();
    }
    #endregion
    #region Elastic mode shapes
    /// <summary>Symmetric mode: Restrainer end ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CrSym(double ke, double κg)
    {
        return (PI / ke * Sin(PI / ke) - κg * Cos(PI / ke)) /
                (PI / ke * Sin(PI / ke) - κg * Cos(PI / ke) + κg);
    }
    /// <summary>Symmetric mode: Gusset ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CgSym(double ke, double κg)
    {
        return -(κg) /
                  (PI / ke * Sin(PI / ke) - κg * Cos(PI / ke) + κg);
    }
    /// <summary>Anti-symmetric mode: Restrainer end ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CrAsym(double ke, double κg, double ξ)
    {
        return ((Pow(PI / ke, 2) + 2 * ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke)) /
         ((Pow(PI / ke, 2) + 2 * ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke) + κg * PI / ke * (1 - 2 * ξ));
    }
    /// <summary>Anti-symmetric mode: Gusset ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CgAsym(double ke, double κg, double ξ)
    {
        return -(κg * PI / ke) /
                 ((Pow(PI / ke, 2) + 2 * ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke) + κg * PI / ke * (1 - 2 * ξ));
    }
    /// <summary>One-sided mode: Restrainer end ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CrOne(double ke, double κg, double ξ)
    {
        return ((Pow(PI / ke, 2) +  ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke)) /
         ((Pow(PI / ke, 2) +  ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke) + κg * PI / ke * (1 -  ξ));
    }
    /// <summary>One-sided mode: Gusset ROTATION / restrainer end lateral DISPLACEMENT</summary>
    internal static double CgOne(double ke, double κg, double ξ)
    {
        return -(κg * PI / ke) /
                 ((Pow(PI / ke, 2) +  ξ * κg) * Sin(PI / ke) - κg * PI / ke * Cos(PI / ke) + κg * PI / ke * (1 -  ξ));
    }
    #endregion
}