using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;
public static class CodeBucklingFunctions
{
    [ExcelFunction(Description = @"Slenderness
        λ=le/ry=π√EA/Ncr")]
    public static double Ncr_slenderness(
         [ExcelArgument("Ncr")] double Ncr,
         [ExcelArgument("Area")] double A,
         [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);

        return PI * Sqrt(E*A / Ncr);
    }

    [ExcelFunction(Description = @"BS5950 Annex C Perry Robertson inelastic column buckling
        Ncr=A·fcr
        where fcr=fe·fy/(φ+√(φ²-fe·fy)), φ=0.5(fy+(1+η)fe)
        fe=π²E/λ², η=a/1000·(λ-0.2·π·√E/fy)>0")]
    public static double Ncr_BS5950(
         [ExcelArgument("Area")] double A,
         [ExcelArgument("Slenderness (Le/r)")] double λ,
         [ExcelArgument("Strut curve (a,b,c,d)")] string curve,
         [ExcelArgument("Yield strength")] double fy,
         [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);

        // BS5950 Tbl strut curve
        double a = 0;
        switch (curve)
        {
            case "a": a = 2.0; break;
            case "b": a = 3.5; break;
            case "c": a = 5.5; break;
            case "d": a = 8.0; break;
            default: return double.NaN;
        }

        // Robertson factor (normalized imperfection)
        double η = Max(0, a / 1000 * (λ - 0.2 * Sqrt(E / fy) * PI));
        // Euler buckling stress
        double fe = Pow(PI, 2) * E / Pow(λ, 2);
        // Perry factor
        double φ = 0.5 * (fy + (η + 1) * fe);
        // inelastic buckling stress
        double fcr = fe * fy / (φ + Sqrt(Pow(φ, 2) - fe * fy));

        return A * fcr;
    }
    [ExcelFunction(Description = @"BS5950 Equivalent imperfection
        δ/L=α·r/c/1000")]
    public static double Ncr_BS5950_δ0(
        [ExcelArgument("Radius of gyration (r) divided by dist from extreme fibre to neutral axis (c)", Name = "r/c")] double r_c,
        [ExcelArgument("Strut curve {a,b,c,d}")] string curve,
        [ExcelArgument("Yield strength")] double fy,
        [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);

        double a = 0;
        switch (curve)
        {
            case "a": a = 2.0; break;
            case "b": a = 3.5; break;
            case "c": a = 5.5; break;
            case "d": a = 8.0; break;
            default: return double.NaN;
        }

        return a * r_c / 1000;
    }

    [ExcelFunction(Description = @"Eurocode 3-1-1 cl6.3.1.2 Perry Robertson inelastic column buckling
        Ncr=A·χ·fy
        where χ=1/(φ+√(φ²-λn²)), φ=0.5(1+η+λn²)
        fe=π²E/λ², λn=λ/π·√fy/E, η=α(λn-0.2)>0")]
    public static double Ncr_EC3(
        [ExcelArgument("Effective Area")] double A,
        [ExcelArgument("Slenderness (Le/r)")] double λ,
        [ExcelArgument("Strut curve (a0,a,b,c,d)")] string curve,
        [ExcelArgument("Yield strength")] double fy,
        [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);
        // EC3 Tbl strut curve
        double α = 0;
        switch (curve)
        {
            case "a0": α = 0.13; break;
            case "a": α = 0.21; break;
            case "b": α = 0.34; break;
            case "c": α = 0.49; break;
            case "d": α = 0.76; break;
            default: return double.NaN;
        }

        // Normalized slenderness
        double λn = λ / Sqrt(E / fy) / PI;
        // Robertson factor (normalized imperfection)
        double η = Max(0, α * (λn - 0.2));
        // Euler buckling stress
        //double fe = Pow(PI, 2) * E / Pow(λ, 2);
        // Perry factor
        double φ = 0.5 * (1 + η + Pow(λn, 2));
        // buckling stress reduction factor
        double χ = 1 / (φ + Sqrt(Pow(φ, 2) - Pow(λn, 2)));

        return A * χ * fy;
    }
    [ExcelFunction(Description = @"Eurocode 3-1-1 Equivalent imperfection
        δ/L=α/π·r/c")]
    public static double Ncr_EC3_δ0(
         [ExcelArgument("Radius of gyration (r) divided by dist from extreme fibre to neutral axis (c)", Name = "r/c")] double r_c,
         [ExcelArgument("Strut curve {a0,a,b,c,d}")] string curve,
         [ExcelArgument("Yield strength")] double fy,
         [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);
        // EC3 Tbl strut curve
        double α = 0;
        switch (curve)
        {
            case "a0": α = 0.13; break;
            case "a": α = 0.21; break;
            case "b": α = 0.34; break;
            case "c": α = 0.49; break;
            case "d": α = 0.76; break;
            default: return double.NaN;
        }

        return α * PI * r_c;
    }

    [ExcelFunction(Description = @"NZS3404 6.2 Perry Robertson/SSRC inelastic col. buckling
        Ncr=A·αc·fy
        where λn=λ·√kf·√fy/250, αa=2100(λn-13.5)/(λn²-15.3λn+2050)
        λe=λn+αa*αb, η=0.00326*(λe-13.5)>0, ξ=((λe/90)²+1+η)/(2·(λe/90)²)
        αc=ξ(1-√(1-(90/λe/ξ)²)")]
    public static double Ncr_NZS3404(
        [ExcelArgument("Net Area")] double A,
        [ExcelArgument("Slenderness (Le/r)")] double λ,
        [ExcelArgument("Form factor (Effecive Area / Gross area)")] double kf,
        [ExcelArgument("Imperfection factor = -1.0~1.0")] double αb,
        [ExcelArgument("Yield strength [MPa]")] double fy,
        [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);
        // Normalized slenderness
        double λn = λ * Sqrt(kf) * Sqrt(fy / 250);
        // Fitted to SSRC curves
        double αa = 2100 * (λn - 13.5) / (Pow(λn, 2) - 15.3 * λn + 2050);
        double λe = λn + αa * αb;
        // Robertson factor (normalized imperfection)
        double η = Max(0, 0.00326 * (λe - 13.5));
        // Perry factor
        double ξ = (Pow(λe / 90, 2) + 1 + η) / (2 * Pow(λe / 90, 2));

        // buckling stress reduction factor
        double αc = ξ * (1 - Sqrt(1 - Pow(90 / λe / ξ, 2)));

        return A * αc * fy;
    }

    [ExcelFunction(Description = @"AISC360 E3 inelastic column buckling
        equivalent to BS5950 strut curve a / EC3 strut curve a0~a / NZS αb = -0.7
        λ < 4.71√fy/E: Ncr=A·fy·0.658^√fy/fe
        λ > 4.71√fy/E: Ncr=A·0.877·fe
        where fe=π²E/λ²")]
    public static double Ncr_AISC360(
        [ExcelArgument("Gross Area")] double A,
        [ExcelArgument("Slenderness (Le/r)")] double λ,
        [ExcelArgument("Yield strength")] double fy,
        [ExcelArgument("Youngs modulus [default=205000MPa")] object _E)
    {
        double E = Check<double>(_E, 205000);
        // Euler buckling stress
        double fe = Pow(PI, 2) * E / Pow(λ, 2);

        if (λ < 4.71 * Sqrt(E / fy))
        {
            return A * Pow(0.658, fy / fe) * fy;
        }
        else
        {
            return A * 0.877 * fe;
        }
    }
}