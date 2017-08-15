using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;

public static class MemberStiffnessFunctions
{

    [ExcelFunction(Description = @"Flexural stiffness
                                    Roarke Tbl 8.1.3c&e
                                    K=3EI/L (pinned-pinned), K=4EI/l(pinned-fixed)")]
    public static double FlexuralStiffness([ExcelArgument("Stiffness")] double I, [ExcelArgument("Span/height")] double L,
         [ExcelArgument("Boundary condition (far end) [default=false]", Name = "isFixed")] object _isFixed, 
         [ExcelArgument("Youngs modulus [default=205000MPa]", Name = "E")] object _E)
    {
        bool isFixed = Check<bool>(_isFixed, false);
        double E = Check<double>(_E, 205000);

        if (I == 0 || E == 0) return 0;
        else return (isFixed ? 4 : 3) * E * I / L;
    }
    [ExcelFunction(Description = @"Torsional stiffness
                                    Lt/L=0.5,warping fixed: Roarke 10.3.1g
                                    Lt/L=0.5,warping released: Roarke 10.3.1e
                                    general,warping fixed: SCI P385 tbl C-1.4(typos), AISC DG9 Case 6 p.110
                                    general,warping released: SCI P385 tbl C-1.3, AISC DG9 Case 3 p.110")]
    public static double TorsionalStiffness([ExcelArgument("Torional constant")] double J, [ExcelArgument("Warping constant")] double Cw,
        [ExcelArgument("Length to applied load")] double Lt, [ExcelArgument("Full span")] double L,
          [ExcelArgument("Boundary conditions [default=false]", Name = "isWarpingRestrained")] object _isWarpingRestrained,
          [ExcelArgument("Youngs modulus [default=205000MPa]", Name = "E")] object _E,
          [ExcelArgument("Shear modulus [default=205000MPa/2(1+0.3)]", Name = "G")] object _G)
    {
        bool isWarpingRestrained = Check<bool>(_isWarpingRestrained, false);
        double E = Check<double>(_E, 205000);
        double G = Check<double>(_G, 205000/2/(1+0.3));

        if (J == 0 || Cw == 0 || E == 0 || G == 0)
        {
            return 0;
        }
        else if (Lt / L == 0.5)
        {
            double v = Sqrt(G * J / (E * Cw));

            if (isWarpingRestrained)
            {
                // Roarke 10.3.1g
                return G * J * v / (Lt * v / 2 - Tanh(Lt * v / 2));

                // == Takeuchi EESD 2016 Eqn C2
                // return 2 * G * J * v / ((Cosh(Lg * v) - 1) ^ 2 / Sinh(Lg * v) - Sinh(Lg * v) + Lg * v)
            }
            else
            {
                // Roarke 10.3.1e (=2*Kfixed(2Lg))
                return 2 * G * J * v / (Lt * v - Tanh(Lt * v));
            }
        }
        else
        {
            double v = Sqrt(G * J / (E * Cw));
            if (isWarpingRestrained)
            {
                double K1 = ((1.0 - Cosh(Lt * v)) / Tanh(L * v) + (Cosh(Lt * v) - 1.0) / Sinh(L * v) + Sinh(Lt * v) - Lt * v)
                            / ((Cosh(L * v) + Cosh(Lt * v) * Cosh(L * v) - Cosh(Lt * v) - 1) / Sinh(L * v)
                            - Sinh(Lt * v) + (Lt / L - 1) * L * v);
                double K3 = 1.0 / Sinh(L * v) + Sinh(Lt * v) - Cosh(Lt * v) / Tanh(L * v);
                double K4 = Sinh(Lt * v) - Cosh(Lt * v) / Tanh(L * v) + 1.0 / Tanh(L * v);

                //SCI P385 Table C-1.4: typos, use SCI P057 / AISC DG9 Case 6 p.110
                return (K1 + 1) * G * J * v / ((K1 * K3 + K4) * (Cosh(Lt * v) - 1.0) - Sinh(Lt * v) + Lt * v);
            }
            else
            {
                //SCI P385 Table C-1.3 / AISC DG9 Case 3 p.110
                return G * J * v / ((1 - Lt / L) * Lt * v
                    + (Sinh(Lt * v) / Tanh(L * v) - Cosh(Lt * v)) * Sinh(Lt * v));
            }
        }
    }

    [ExcelFunction(Description = @"Transformed rotational stiffness
                                  Kf=1/(sin²θ/Kxx+cos²θ/Kxx)")]
    public static object TransformedStiffness(
        [ExcelArgument("Stiffness about horizontal axis")] object _Kxx,
        [ExcelArgument("Stiffness about vertical axis")] object _Kzz, 
        [ExcelArgument("BRB angle (rad) from the horizontal")] double theta)
    {
        double Kxx = ToDoubleInfinity(_Kxx);
        double Kzz = ToDoubleInfinity(_Kzz);

        //theta in radians
        if (double.IsInfinity(Kxx) && double.IsInfinity(Kzz)) return _Kxx;
        else if (double.IsInfinity(Kxx)) return Kzz / Pow(Cos(theta), 2);
        else if (double.IsInfinity(Kzz)) return Kxx / Pow(Sin(theta), 2);
        else return 1 / (Pow(Sin(theta), 2) / Kxx + Pow(Cos(theta), 2) / Kzz);
    }
    [ExcelFunction(Description = @"Effective connection length
                                  L'=(Lg²/Kg+Lf²/Kf)/(Lg/Kg+Lf/Kf)")]
    public static double BRB_EquivConn_Lc(
        [ExcelArgument("Length of gusset")] double Lg, 
        [ExcelArgument("Length of gusset+framing")] double Lf,
        [ExcelArgument("Gusset rotational stiffness (use ∞ if rigid)")] object _Kg,
        [ExcelArgument("Framing rotational stiffness (use ∞ if rigid)")] object _Kf)
    {
        double Kf = ToDoubleInfinity(_Kf);
        double Kg = ToDoubleInfinity(_Kg);

        if (double.IsInfinity(Kg) || Kf==0) return Lf;
        else if (double.IsInfinity(Kf) || Kg == 0) return Lg;
        else return (Pow(Lg, 2) / Kg + Pow(Lf, 2) / Kf) / (Lg / Kg + Lf / Kf);
    }
    [ExcelFunction(Description = @"Effective connection length
                                  L'=(Lg²/Kg+Lf²/Kf)/(Lg/Kg+Lf/Kf)")]
    public static object BRB_EquivConn_Kc(
        [ExcelArgument("Effective length")] double Leq,
        [ExcelArgument("Length of gusset")] double Lg,
        [ExcelArgument("Length of gusset+framing")] double Lf,
        [ExcelArgument("Gusset rotational stiffness (use ∞ if rigid)")] object _Kg,
        [ExcelArgument("Framing rotational stiffness (use ∞ if rigid)")] object _Kf)
    {
        double Kf = ToDoubleInfinity(_Kf);
        double Kg = ToDoubleInfinity(_Kg);
        if (Kg == 0 || Kf == 0) return 0;
        else if (double.IsInfinity(Kg) && double.IsInfinity(Kf)) return _Kg;
        else if (double.IsInfinity(Kg)) return Kf;
        else if (double.IsInfinity(Kf)) return Kg;
        else return 1 / (Lg / Leq / Kg + Lf / Leq / Kf);
    }    
}