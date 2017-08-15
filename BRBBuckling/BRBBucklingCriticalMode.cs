using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;
using static BRBBucklingFunctions;
using System;

public static class BRBBucklingCriticalMode
{
    [ExcelFunction(Description = @"Determine critical buckling mode shape (Sym vs Asym)")]
    public static double BRBNcr_Sitler2CritKg(
        [ExcelArgument(Name = "κg_min", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κgmin,
        [ExcelArgument(Name = "κg_max", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κgmax,
        [ExcelArgument(Name = "κg_tol", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κgtol,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
        [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end")] double Δr,
        [ExcelArgument(Name = "θpr", Description = "Rest end plastic moment capacity (reduced for N)")] object _θpr,
        [ExcelArgument(Name = "θpg", Description = "Gusset plastic moment capacity (reduced for N)")] object _θpg,
        [ExcelArgument(Name = "θ0r", Description = "Moment at rest end due to out-of-plane drift")] object _θ0r,
        [ExcelArgument(Name = "θ0g", Description = "Moment at gusset due to out-of-plane drift")] object _θ0g)
    {
        double θpr = Check<double>(_θpr, 0);
        double θpg = Check<double>(_θpg, 0);
        double θ0r = Check<double>(_θ0r, 0);
        double θ0g = Check<double>(_θ0g, 0);

        Func<double, double> ModeRatio = (κg) =>
        {
            double keSym = BRBke_sym(κg, κr, ξ, γ, ExcelMissing.Value,ExcelMissing.Value, ExcelMissing.Value);
            double crSym = BRBcrSym(keSym, κg, true);
            double cgSym = BRBcgSym(keSym, κg, true);

            double keAsym = BRBke_asym(κg, κr, ξ, γ, ExcelMissing.Value, ExcelMissing.Value, ExcelMissing.Value);
            double crAsym = BRBcrAsym(keAsym, κg, ξ, true);
            double cgAsym = BRBcgAsym(keAsym, κg, ξ, true);

            double NcrrSym = BRBNcr_Sitler2(keSym, crSym, κr, θpr, θ0r, Δr);
            double NcrrAsym = BRBNcr_Sitler2(keAsym, crAsym, κr, θpr, θ0r, Δr);
            double NcrgSym = BRBNcr_Sitler2(keSym, cgSym, κg, θpg, θ0g, Δr);
            double NcrgAsym = BRBNcr_Sitler2(keAsym, cgAsym, κg, θpg, θ0g, Δr);

            if (θpr == 0) return NcrgSym / NcrgAsym ;
            else if (θpg == 0) return NcrrSym / NcrrAsym ;
            else return Min(NcrgSym, NcrrSym) / Min(NcrgAsym, NcrrAsym);
        };

        return BRBBuckling.NLEquationSolver.BruteForce(ModeRatio,1.0, κgmin, κgmax, κgtol);
    }
    [ExcelFunction(Description = @"Determine critical buckling mode shape (Sym vs Asym)")]
    public static double BRBNcr_Sitler2CritKr(
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr_min", Description = "Normalized restrainer end stiffness (Kr·ξL0/γEIr)")] double κrmin,
        [ExcelArgument(Name = "κr_max", Description = "Normalized restrainer end stiffness (Kr·ξL0/γEIr)")] double κrmax,
        [ExcelArgument(Name = "κr_tol", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κrtol,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
        [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end")] double Δr,
        [ExcelArgument(Name = "θpr", Description = "Rest end plastic moment capacity (reduced for N)")] object _θpr,
        [ExcelArgument(Name = "θpg", Description = "Gusset plastic moment capacity (reduced for N)")] object _θpg,
        [ExcelArgument(Name = "θ0r", Description = "Moment at rest end due to out-of-plane drift")] object _θ0r,
        [ExcelArgument(Name = "θ0g", Description = "Moment at gusset due to out-of-plane drift")] object _θ0g)
    {
        double θpr = Check<double>(_θpr, 0);
        double θpg = Check<double>(_θpg, 0);
        double θ0r = Check<double>(_θ0r, 0);
        double θ0g = Check<double>(_θ0g, 0);

        Func<double, double> ModeRatio = (κr) =>
        {
            double keSym = BRBke_sym(κg, κr, ξ, γ, ExcelMissing.Value, ExcelMissing.Value, ExcelMissing.Value);
            double crSym = BRBcrSym(keSym, κg, true);
            double cgSym = BRBcgSym(keSym, κg, true);

            double keAsym = BRBke_asym(κg, κr, ξ, γ, ExcelMissing.Value, ExcelMissing.Value, ExcelMissing.Value);
            double crAsym = BRBcrAsym(keAsym, κg, ξ, true);
            double cgAsym = BRBcgAsym(keAsym, κg, ξ, true);

            double NcrrSym = BRBNcr_Sitler2(keSym, crSym, κr, θpr, θ0r, Δr);
            double NcrrAsym = BRBNcr_Sitler2(keAsym, crAsym, κr, θpr, θ0r, Δr);
            double NcrgSym = BRBNcr_Sitler2(keSym, cgSym, κg, θpg, θ0g, Δr);
            double NcrgAsym = BRBNcr_Sitler2(keAsym, cgAsym, κg, θpg, θ0g, Δr);

            if (θpr == 0) return NcrgSym / NcrgAsym;
            else if (θpg == 0) return NcrrSym / NcrrAsym;
            else return Min(NcrgSym, NcrrSym) / Min(NcrgAsym, NcrrAsym);
        };

        return BRBBuckling.NLEquationSolver.BruteForce(ModeRatio, 1.0, κrmin, κrmax, κrtol);
    }
}