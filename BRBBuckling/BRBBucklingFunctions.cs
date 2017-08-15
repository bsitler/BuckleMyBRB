using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;

public static class BRBBucklingFunctions
{

    #region Stability limits
    [ExcelFunction(Description = @"Nakamura's BRB buckling method
        Ncr=π^2EI/2L^2
        Tsai, Nakamura et al. 2002: Experimental test of large scale BRB, Passive Control Symposium, CUEE@TokyoTech 
        Bruneau, Uang, Sabelli 2009: Ductile design of steel structures")]
    public static double BRBNcr_Nakamura(
        [ExcelArgument(Name = "L", Description = "Workpoint to restrainer end length")] double L,
        [ExcelArgument(Name = "I", Description = "Neck section modulus")] double I,
        [ExcelArgument(Name = "E", Description = "Young's modulus")] double E)
    {
        return Pow(PI, 2) * I * E / Pow(2 * L, 2);
    }

    [ExcelFunction(Description = @"Okazaki's BRB buckling method
        Ncrᵣ=Kr(1-2ξ)/ξL0(1-ξ)
        Hikino, Okazaki et al. 2012: Out-of-plane stability of BRBs placed in a chevron arrangement, ASCE 139(11)")]
    public static double BRBNcr_Okazaki(
        [ExcelArgument(Name = "ξL0", Description = "Connection length")] double ξL0,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to overall length")] double ξ,
        [ExcelArgument(Name = "Kf", Description = "Rotational stiffness of adjacent framing")] double Kf)
    {
        return Kf / ξL0 * (1 - 2 * ξ) / (1 - ξ);
    }

    [ExcelFunction(Description = @"Takeuchi's BRB buckling method
        Nlim1=((Mʳp-Mʳ0₀)/ar+Nʳcr)/((Mʳp-Mʳ0)/arNᴮcr+1)
        Nlim2=((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/ar)/((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi(
    [ExcelArgument(Name = "Nᴮcr", Description = "Elastic buckling capacity")] double NcrB,
    [ExcelArgument(Name = "Nʳcr", Description = "Inelastic buckling capacity with pinned restrainer ends")] double Ncrr,
    [ExcelArgument(Name = "Mʳp", Description = "Rest or neck plastic moment capacity (reduced for N)")] double Mpr,
    [ExcelArgument(Name = "Mʳ0", Description = "Moment at restrainer end due to out-of-plane drift")] double M0r,
    [ExcelArgument(Name = "Mᵍp", Description = "Gusset plastic moment capacity (reduced for N)")] double Mpg,
    [ExcelArgument(Name = "Mᵍ0", Description = "Moment at gusset due to out-of-plane drift")] double M0g,
    [ExcelArgument(Name = "ar", Description = "Imperfection at restrainer end")] double ar)
    {
        return Min(
            BRBNcr_Takeuchi_Nlim1(NcrB, Ncrr, Mpr, M0r, ar),
            BRBNcr_Takeuchi_Nlim2(NcrB, Mpr, M0r, Mpg, M0g, ar));
    }
    [ExcelFunction(Description = @"Takeuchi's BRB buckling method
        Nlim2=((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/ar)/((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi_Nlim2(
    [ExcelArgument(Name = "Nᴮcr", Description = "Elastic buckling capacity")] double NcrB,
    [ExcelArgument(Name = "Mʳp", Description = "Rest or neck plastic moment capacity (reduced for N)")] double Mpr,
    [ExcelArgument(Name = "Mʳ0", Description = "Moment at restrainer end due to out-of-plane drift")] double M0r,
    [ExcelArgument(Name = "Mᵍp", Description = "Gusset plastic moment capacity (reduced for N)")] double Mpg,
    [ExcelArgument(Name = "Mᵍ0", Description = "Moment at gusset due to out-of-plane drift")] double M0g,
    [ExcelArgument(Name = "ar", Description = "Imperfection at restrainer end")] double ar)
    {
        return (Max(0, Mpr - M0r) + Max(0, Mpg - M0g)) / ar / ((Max(0, Mpr - M0r) + Max(0, Mpg - M0g)) / ar / NcrB + 1);
    }
    [ExcelFunction(Description = @"Takeuchi's BRB buckling method
        Nlim1=((Mʳp-Mʳ0₀)/ar+Nʳcr)/((Mʳp-Mʳ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi_Nlim1(
    [ExcelArgument(Name = "Nᴮcr", Description = "Elastic buckling capacity")] double NcrB,
    [ExcelArgument(Name = "Nʳcr", Description = "Inelastic buckling capacity with pinned restrainer ends")] double Ncrr,
    [ExcelArgument(Name = "Mʳp", Description = "Rest or neck plastic moment capacity (reduced for N)")] double Mpr,
    [ExcelArgument(Name = "Mʳ0", Description = "Moment at restrainer end due to out-of-plane drift")] double M0r,
    [ExcelArgument(Name = "ar", Description = "Imperfection at restrainer end")] double ar)
    {
        return (Max(0, Mpr - M0r) / ar + Ncrr) / (Max(0, Mpr - M0r) / ar / NcrB + 1);
    }

    [ExcelFunction(Description = @"Modified Takeuchi's BRB buckling method
        Ncr=((Mp-M0)/c/ar)/((Mp-M0)/c/ar/Nᴮcr+1)")]
    public static double BRBNcr_Sitler(
    [ExcelArgument(Name = "Nᴮcr", Description = "Elastic buckling capacity")] double NcrB,
    [ExcelArgument(Name = "Mp", Description = "Hinge plastic moment capacity (reduced for N)")] double Mp,
    [ExcelArgument(Name = "M0", Description = "Moment at hinge due to out-of-plane drift")] double M0,
    [ExcelArgument(Name = "c", Description = "Moment modification factor (Mpδ=c·N·yr)")] double c,
    [ExcelArgument(Name = "ar", Description = "Imperfection at restrainer end")] double ar)
    {
        return ((Mp - M0) / c / ar) / ((Mp - M0) / c / ar / NcrB + 1);
    }
    #endregion
    #region Normalized Stability limits
    [ExcelFunction(Description = @"Nakamura's BRB buckling method (normalized)
        Ncr/(π²γEIr/(2ξL0)²)=η/2²ς²
        Tsai, Nakamura et al. 2002: Experimental test of large scale BRB, Passive Control Symposium, CUEE@TokyoTech 
        Bruneau, Uang, Sabelli 2009: Ductile design of steel structures")]
    public static double BRBNcr_Nakamura2(
        [ExcelArgument(Name = "ς", Description = "Ratio of connection workpoint to connection rotation lengths (ξwp/ξ)")] double ς,
        [ExcelArgument(Name = "η", Description = "Ratio of neck to connection stiffness (EIn/γEIr)")] double η)
    {
        return η / Pow(2 * ς, 2);
    }

    [ExcelFunction(Description = @"AIJ 2014 BRB buckling method (normalized)
        Ncr/(π²γEIr/(ξL0)²)=θᵍy·κg(1-2ξ)/(θᵍy·κg+π²Δr(1-2ξ))
        Tsai, Nakamura et al. 2002: Experimental test of large scale BRB, Passive Control Symposium, CUEE@TokyoTech 
        Bruneau, Uang, Sabelli 2009: Ductile design of steel structures")]
    public static double BRBNcr_AIJ2(
          [ExcelArgument(Name = "η", Description = "Ratio of neck to connection stiffness (EIn/γEIr)")] double η,
          [ExcelArgument(Name = "ξ", Description = "Ratio of connection to overall length")] double ξ,
          [ExcelArgument(Name = "ξg", Description = "Ratio of connection to overall length")] double ξg,
          [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
          [ExcelArgument(Name = "θᵍy", Description = "Gusset yield rotation (Mᵍy/Kg)")] double θyg,
          [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end (ar/ξjL0)")] double Δr)
    {
        return θyg * κg * (1 - 2 * ξg) * η / (θyg * κg * Pow(2 * ξg / ξ, 2) + Pow(PI, 2) * (1 - 2 * ξg) * η * Δr);
    }
    [ExcelFunction(Description = @"Okazaki's BRB buckling method (normalized)
        Ncr/(π²γEIr/(ξL0)²)=κf/π²(1-2ξ)/(1-ξ)
        Hikino, Okazaki et al. 2012: Out-of-plane stability of BRBs placed in a chevron arrangement, ASCE 139(11)")]
    public static double BRBNcr_Okazaki2(
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to overall length")] double ξ,
        [ExcelArgument(Name = "κf", Description = "Normalized rotational stiffness of adjcent framing (Kf)")] double κf)
    {
        return κf / Pow(PI, 2) * (1 - 2 * ξ) / (1 - ξ);
    }

    [ExcelFunction(Description = @"Takeuchi's BRB buckling method (normalized)
        Nlim1=((Mʳp-Mʳ0₀)/ar+Nʳcr)/((Mʳp-Mʳ0)/arNᴮcr+1)
        Nlim2=((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/ar)/((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi2(
        [ExcelArgument(Name = "keᴮcr", Description = "Elastic buckling length factor (1/√(Nᴮcr/N⁰cr))")] double keB,
        [ExcelArgument(Name = "keʳcr", Description = "Inelastic buckling length factor with pinned restrainer ends (1/√(Nʳcr/N⁰cr))")] double ker,
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "θʳp", Description = "Rest or neck plastic moment capacity (Mʳp/Kr)")] double θpr,
        [ExcelArgument(Name = "θʳ0", Description = "Moment at restrainer end due to out-of-plane drift (Mʳ0/Kr)")] double θ0r,
        [ExcelArgument(Name = "θᵍp", Description = "Gusset plastic moment capacity (Mᵍp/Kg)")] double θpg,
        [ExcelArgument(Name = "θᵍ0", Description = "Moment at gusset due to out-of-plane drift (Mᵍ0/Kg)")] double θ0g,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end (ar/ξL0)")] double Δr)
    {
        return Min(
            BRBNcr_Takeuchi2_Nlim1(keB, ker, κg, κr, θpr, θ0r, Δr),
            BRBNcr_Takeuchi2_Nlim2(keB, κg, κr, θpr, θ0r, θpg, θ0g, Δr));
    }
    [ExcelFunction(Description = @"Takeuchi's BRB buckling method Nlim2 (normalized)
        Nlim1=((Mʳp-Mʳ0₀)/ar+Nʳcr)/((Mʳp-Mʳ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi2_Nlim1(
        [ExcelArgument(Name = "keᴮcr", Description = "Elastic buckling length factor (1/√(Nᴮcr/N⁰cr))")] double keB,
        [ExcelArgument(Name = "keʳcr", Description = "Inelastic buckling length factor with pinned restrainer ends (1/√(Nʳcr/N⁰cr))")] double ker,
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "θʳp", Description = "Rest or neck plastic moment capacity (Mʳp/Kr)")] double θpr,
        [ExcelArgument(Name = "θʳ0", Description = "Moment at restrainer end due to out-of-plane drift (Mʳ0/Kr)")] double θ0r,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end (ar/ξL0)")] double Δr)
    {
        return ((Max(0, θpr - θ0r) * κr) / Pow(PI, 2) / Δr + 1 / Pow(ker, 2)) /
            ((Max(0, θpr - θ0r) * κr) / Pow(PI, 2) / Δr * Pow(keB, 2) + 1);
    }
    [ExcelFunction(Description = @"Takeuchi's BRB buckling method Nlim1 (normalized)
        Nlim2=((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/ar)/((Mʳp-Mʳ0+Mᵍp-Mᵍ0)/arNᴮcr+1)
        Takeuchi et al 2014: Out-of-plane stability of BRB incl moment transfer capacity. EESD 43(6))")]
    public static double BRBNcr_Takeuchi2_Nlim2(
        [ExcelArgument(Name = "keᴮcr", Description = "Elastic buckling length factor (1/√(Nᴮcr/N⁰cr))")] double keB,
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "θʳp", Description = "Rest or neck plastic moment capacity (Mʳp/Kr)")] double θpr,
        [ExcelArgument(Name = "θʳ0", Description = "Moment at restrainer end due to out-of-plane drift (Mʳ0/Kr)")] double θ0r,
        [ExcelArgument(Name = "θᵍp", Description = "Gusset plastic moment capacity (Mᵍp/Kg)")] double θpg,
        [ExcelArgument(Name = "θᵍ0", Description = "Moment at gusset due to out-of-plane drift (Mᵍ0/Kg)")] double θ0g,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end (ar/ξL0)")] double Δr)
    {
        return (Max(0, θpr - θ0r) * κr + Max(0, θpg - θ0g) * κg) / Pow(PI, 2) / Δr /
            ((Max(0, θpr - θ0r) * κr + Max(0, θpg - θ0g) * κg) / Pow(PI, 2) / Δr * Pow(keB, 2) + 1);
    }

    [ExcelFunction(Description = @"Modified Takeuchi's BRB buckling method (normalized)
        Takeuchi et al 2014: Out-of-plane stability of BRB including moment transfer capacity. Earth Eng Strut Dyn 43(6))")]
    public static double BRBNcr_Sitler2(
        [ExcelArgument(Name = "keᴮcr", Description = "Elastic buckling length factor (1/√(Nᴮcr/N⁰cr))")] double keB,
        [ExcelArgument(Name = "c", Description = "Pδ moment factor (MPδ=c·N·yr)")] double c,
        [ExcelArgument(Name = "κ", Description = "Normalized gusset rotational stiffness (K·ξL0/γEIr)")] double κ,
        [ExcelArgument(Name = "θp", Description = "Rest or neck plastic moment capacity (Mp/K)")] double θp,
        [ExcelArgument(Name = "θ0", Description = "Moment at restrainer end due to out-of-plane drift (M0/K)")] double θ0,
        [ExcelArgument(Name = "Δr", Description = "Imperfection at restrainer end (ar/ξL0)")] double Δr)
    {
        return ((Max(0, θp - θ0) * κ) / Pow(PI, 2) / c / Δr) /
            ((Max(0, θp - θ0) * κ) / Pow(PI, 2) / c / Δr * Pow(keB, 2) + 1);
    }
    [ExcelFunction(Description = @"Modified Takeuchi's BRB buckling method
        Ncr=((Mp-M0)/c/ar)/((Mp-M0)/c/ar/Nᴮcr+1)")]
    public static double BRBNcr_Sitler2Direct(
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
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

        if (θpr == 0) return Min(NcrgSym, NcrgAsym);
        else if (θpg == 0) return Min(NcrrSym, NcrrAsym);
        else return Min(Min(NcrgSym, NcrgAsym), Min(NcrrSym, NcrrAsym));
    }



    #endregion
    #region Normalization
    [ExcelFunction(Description = @"Normalized rotational stiffness
        κg1=Kg1·γ1EIr/ξ1L0
        κr1=Kr1·γ1EIr/ξ1L0
        κg2=Kg2·γ2EIr/ξ2L0
        κr2=Kr2·γ2EIr/ξ2L0")]
    public static double BRBnorm_K(
     [ExcelArgument(Name = "K", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double K,
     [ExcelArgument(Name = "ξL0", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξL0,
     [ExcelArgument(Name = "γEIr", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γEIr)
    {
        return K * ξL0 / γEIr;
    }
    [ExcelFunction(Description = @"Index buckling load
        Ncr0=π²γ1EIr/(ξ1L0)²")]
    public static double BRBNcr0(
     [ExcelArgument(Name = "ξ1L0", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ1L0,
     [ExcelArgument(Name = "γ1EIr", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ1EIr)
    {
        return Pow(PI, 2) * γ1EIr / Pow(ξ1L0, 2);
    }
    #endregion

    #region Elastic buckling loads
    [ExcelFunction(Description = @"Symmetric mode elastic buckling load
            Kg      Kr @----------------@ Kr      Kg
            @---------/                  \---------@
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBNcre_sym(
         [ExcelArgument(Name = "Kg", Description = "Gusset rotational stiffness")] double Kg,
         [ExcelArgument(Name = "Kr", Description = "Restrainer end rotational stiffness")] double Kr,
         [ExcelArgument(Name = "ξL0", Description = "Connection buckling length")] double ξL0,
         [ExcelArgument(Name = "γEIr", Description = "Connection flexural stiffness")] double γEIr,
         [ExcelArgument(Name = "L0", Description = "Full brace buckling length")] double L0,
         [ExcelArgument(Name = "EIr", Description = "Restrainer flexural stiffness")] double EIr,
         [ExcelArgument(Name = "Tolerance", Description = "Load increment [default = 1·force unit]")] object _tol)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 1);
        double Ncr0 = BRBNcr0(ξL0, γEIr);

        double ke = ElasticBuckling.ke_crit(N => ElasticBuckling.det_sym(N,
            BRBnorm_K(Kg, ξL0, γEIr), BRBnorm_K(Kr, ξL0, γEIr), ξL0 / L0, γEIr / EIr),
            1, tol / Ncr0, 1.0);
        if (double.IsInfinity(ke)) return 0;
        else return Ncr0 / Pow(ke, 2);
    }
    [ExcelFunction(Description = @"Anti-Symmetric mode elastic buckling load
                    -----@-----
            @----/     Kr    \----\     Kr    /----@
            Kg                     -----@-----     Kg
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBNcre_asym(
         [ExcelArgument(Name = "Kg", Description = "Gusset rotational stiffness")] double Kg,
         [ExcelArgument(Name = "Kr", Description = "Restrainer end rotational stiffness")] double Kr,
         [ExcelArgument(Name = "ξL0", Description = "Connection buckling length")] double ξL0,
         [ExcelArgument(Name = "γEIr", Description = "Connection flexural stiffness")] double γEIr,
         [ExcelArgument(Name = "L0", Description = "Full brace buckling length")] double L0,
         [ExcelArgument(Name = "EIr", Description = "Restrainer flexural stiffness")] double EIr,
         [ExcelArgument(Name = "Tolerance", Description = "Load increment [default = 1·force unit]")] object _tol)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 1);
        double Ncr0 = BRBNcr0(ξL0, γEIr);

        double ke = ElasticBuckling.ke_crit(N => ElasticBuckling.det_asym(N,
                BRBnorm_K(Kg, ξL0, γEIr), BRBnorm_K(Kr, ξL0, γEIr), ξL0 / L0, γEIr / EIr),
                1, tol / Ncr0, 1);
        if (double.IsInfinity(ke)) return 0;
        else return Ncr0 / Pow(ke, 2);
    }
    [ExcelFunction(Description = @"One-sided mode elastic buckling load
            Kg   ------@--------
            @---/      Kr       \--------O==========@
         Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBNcre_one(
         [ExcelArgument(Name = "Kg", Description = "Gusset rotational stiffness")] double Kg,
         [ExcelArgument(Name = "Kr", Description = "Restrainer end rotational stiffness")] double Kr,
         [ExcelArgument(Name = "ξL0", Description = "Connection buckling length")] double ξL0,
         [ExcelArgument(Name = "γEIr", Description = "Connection flexural stiffness")] double γEIr,
         [ExcelArgument(Name = "L0", Description = "Full brace buckling length")] double L0,
         [ExcelArgument(Name = "EIr", Description = "Restrainer flexural stiffness")] double EIr,
         [ExcelArgument(Name = "Tolerance", Description = "Load increment [default = 1·force unit]")] object _tol)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 1);
        double Ncr0 = BRBNcr0(ξL0, γEIr);

        double ke = ElasticBuckling.ke_crit(N => ElasticBuckling.det_one(N,
            BRBnorm_K(Kg, ξL0, γEIr), BRBnorm_K(Kr, ξL0, γEIr), ξL0 / L0, γEIr / EIr),
            1, tol / Ncr0, 1);
        if (double.IsInfinity(ke)) return 0;
        else return Ncr0 / Pow(ke, 2);
    }
    [ExcelFunction(Description = @"General chevron config. elastic buckling load
        Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBNcre_chev(
         [ExcelArgument(Name = "modeNum", Description = "Mode number")] double modeNum,
         [ExcelArgument(Name = "Kg1", Description = "Gusset rotational stiffness")] double Kg1,
         [ExcelArgument(Name = "Kg2", Description = "Gusset rotational stiffness")] double Kg2,
         [ExcelArgument(Name = "Kr1", Description = "Restrainer end rotational stiffness")] double Kr1,
         [ExcelArgument(Name = "Kr2", Description = "Restrainer end rotational stiffness")] double Kr2,
         [ExcelArgument(Name = "ξ1L0", Description = "Connection buckling length")] double ξ1L0,
         [ExcelArgument(Name = "ξ2L0", Description = "Connection buckling length")] double ξ2L0,
         [ExcelArgument(Name = "γ1EIr", Description = "Connection flexural stiffness")] double γ1EIr,
         [ExcelArgument(Name = "γ2EIr", Description = "Connection flexural stiffness")] double γ2EIr,
         [ExcelArgument(Name = "L0", Description = "Full brace buckling length")] double L0,
         [ExcelArgument(Name = "EIr", Description = "Restrainer flexural stiffness")] double EIr,
         [ExcelArgument(Name = "Tolerance", Description = "Load increment [default = 1·force unit]")] object _tol)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 1);
        double Ncr0 = BRBNcr0(ξ1L0, γ1EIr);

        double ke = ElasticBuckling.ke_crit(N => ElasticBuckling.det_chev(N,
            BRBnorm_K(Kg1, ξ1L0, γ1EIr), BRBnorm_K(Kg2, ξ2L0, γ2EIr),
            BRBnorm_K(Kr1, ξ1L0, γ1EIr), BRBnorm_K(Kr2, ξ2L0, γ2EIr),
            ξ1L0 / L0, ξ2L0 / L0, γ1EIr / EIr, γ2EIr / EIr),
            modeNum, tol / Ncr0, 1);

        if (double.IsInfinity(ke)) return 0;
        else return Ncr0 / Pow(ke, 2);
    }
    #endregion
    #region Normalized Elastic buckling loads
    [ExcelFunction(Description = @"Symmetric mode effective length factor
            NcrB=π²γEIr/(ke·ξL0)²
            κg      κr @----------------@ κr      κg
            @---------/                  \---------@
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBke_sym(
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
        [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
        [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
        [ExcelArgument(Name = "Mode", Description = "Mode number [default = 1]")] object _mode,
        [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)

    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double mode = Check<double>(_mode, 1);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_sym(N, κg, κr, ξ, γ), mode, tol, max);
    }

    [ExcelFunction(Description = @"Anti-Symmetric mode effective length factor
            NcrB=π²γEIr/(ke·ξL0)²
                  -----@-----
            @----/     κr    \----\     κr    /----@
            κg                     -----@-----     κg
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBke_asym(
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
        [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
        [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
        [ExcelArgument(Name = "Mode", Description = "Mode number [default = 1]")] object _mode,
        [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double mode = Check<double>(_mode, 1);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_asym(N, κg, κr, ξ, γ), mode, tol, max);
    }

    [ExcelFunction(Description = @"One-sided mode effective length factor
            NcrB=π²γEIr/(ke·ξL0)²
            κg   ------@--------
            @---/      κr       \--------O==========@
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBke_one(
            [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
            [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
            [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
            [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
            [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
            [ExcelArgument(Name = "Mode", Description = "Mode number [default = 1]")] object _mode,
            [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double mode = Check<double>(_mode, 1);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_one(N, κg, κr, ξ, γ), mode, tol, max);
    }
    [ExcelFunction(Description = @"General chevron config. effective length factor
            NcrB=π²γ1EIr/(ke·ξ1L0)²
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
    public static double BRBke_chev(
        [ExcelArgument(Name = "modeNum", Description = "Mode number")] double modeNum,
        [ExcelArgument(Name = "κg1", Description = "Normalized gusset rotational stiffness (Kg1·ξ1L0/γ1EIr)")] double κg1,
        [ExcelArgument(Name = "κg2", Description = "Normalized gusset rotational stiffness (Kg2·ξ2L0/γ2EIr)")] double κg2,
        [ExcelArgument(Name = "κr1", Description = "Normalized restrainer end rotational stiffness (Kr1·ξ1L0/γ1EIr)")] double κr1,
        [ExcelArgument(Name = "κr2", Description = "Normalized restrainer end rotational stiffness (Kr2·ξ2L0/γ2EIr)")] double κr2,
        [ExcelArgument(Name = "ξ1", Description = "Ratio of connection to full buckling length (ξ1L0/L0)")] double ξ1,
        [ExcelArgument(Name = "ξ2", Description = "Ratio of connection to full buckling length (ξ2L0/L0)")] double ξ2,
        [ExcelArgument(Name = "γ1", Description = "Ratio of connection to restrainer flexural stiffness (γ1EIr/EIr)")] double γ1,
        [ExcelArgument(Name = "γ2", Description = "Ratio of connection to restrainer flexural stiffness (γ2EIr/EIr)")] double γ2,
        [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
        [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)

    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_chev(N, κg1, κg2, κr1, κr2, ξ1, ξ2, γ1, γ2), modeNum, tol, max);
    }


    [ExcelFunction(Description = @"Elastic buckling mode shape of stick-spring system
        Mpδ = cr·N·yr, where: yr is rest. end lateral displacement, N is applied compression
            Mpδ,yr @----------------@
        @---------/                 \---------@")]
    public static double BRBcrSym(
        [ExcelArgument("Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
        [ExcelArgument("Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
          [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CrSym(ke, κg);
        else return Abs(ElasticBuckling.CrSym(ke, κg));
    }
    [ExcelFunction(Description = @"Elastic buckling mode shape of stick-spring system
        Mpδ = cg·N·yr, where: yr is rest. end lateral displacement, N is applied compression
        Mpδ     yr @----------------@
        @---------/                 \---------@")]
    public static double BRBcgSym(
            [ExcelArgument(Name = "ke", Description = "Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
            [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
            [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CgSym(ke, κg);
        else return Abs(ElasticBuckling.CgSym(ke, κg));
    }
    [ExcelFunction(Description = @"Elastic buckling mode shape of stick-spring system
        Mpδ = cr·N·yr, where: yr is rest. end lateral displacement, N is applied compression
              -----@-----
        @----/  Mpδ,yr   \----\           /----@
                               -----@-----     ")]
    public static double BRBcrAsym(
            [ExcelArgument(Name = "ke", Description = "Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
            [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
            [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
            [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CrAsym(ke, κg, ξ);
        else return Abs(ElasticBuckling.CrAsym(ke, κg, ξ));
    }
    [ExcelFunction(Description = @"Elastic buckling mode shape of stick-spring system
        Mpδ = cg·N·yr, where: yr is rest. end lateral displacement, N is applied compression
        Mpδ   -----@-----
        @----/      yr   \----\           /----@
                               -----@-----     ")]
    public static double BRBcgAsym(
        [ExcelArgument(Name = "ke", Description = "Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
        [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
        [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
          [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CgAsym(ke, κg, ξ);
        else return Abs(ElasticBuckling.CgAsym(ke, κg, ξ));
    }
    public static double BRBcrOne(
          [ExcelArgument(Name = "ke", Description = "Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
          [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
          [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
          [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CrOne(ke, κg, ξ);
        else return Abs(ElasticBuckling.CrOne(ke, κg, ξ));
    }
    [ExcelFunction(Description = @"Elastic buckling mode shape of stick-spring system
        Mpδ = cg·N·yr, where: yr is rest. end lateral displacement, N is applied compression
        Mpδ   -----@-----
        @----/      yr   \----\           /----@
                               -----@-----     ")]
    public static double BRBcgOne(
         [ExcelArgument(Name = "ke", Description = "Effective length factor (NcrB=π²γEIr/(ke·ξL0)²)")] double ke,
         [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
         [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
         [ExcelArgument("Return absolute value [default = true]", Name = "absolute")] object _absolute)
    {
        bool absolute = Check<bool>(_absolute, true);

        if (!absolute) return ElasticBuckling.CgOne(ke, κg, ξ);
        else return Abs(ElasticBuckling.CgOne(ke, κg, ξ));
    }
    #endregion
    [ExcelFunction("get excel error")]
    public static double GetError()
    {
        return (double)ExcelDna.Integration.ExcelError.ExcelErrorValue;
    }

    [ExcelFunction(Description = @"Symmetric mode effective length factor
            Simplified buckling mode excluding restrainer Pδ contribution
            @----------@@
            κg         κr'")]
    public static double BRBke_sym2(
          [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
          [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
          [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
          [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
          [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
          [ExcelArgument(Name = "Mode", Description = "Mode number [default = 1]")] object _mode,
          [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double mode = Check<double>(_mode, 1);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_sym_connection(N, κg, κr, ξ, γ), mode, tol, max);
    }

    [ExcelFunction(Description = @"Anti-Symmetric mode effective length factor
            Simplified buckling mode excluding restrainer Pδ contribution
                    -----@@
            @----/     κr'
            κg")]
    public static double BRBke_asym2(
          [ExcelArgument(Name = "κg", Description = "Normalized gusset rotational stiffness (Kg·ξL0/γEIr)")] double κg,
          [ExcelArgument(Name = "κr", Description = "Normalized restrainer end rotational stiffness (Kr·ξL0/γEIr)")] double κr,
          [ExcelArgument(Name = "ξ", Description = "Ratio of connection to full buckling length (ξL0/L0)")] double ξ,
          [ExcelArgument(Name = "γ", Description = "Ratio of connection to restrainer flexural stiffness (γEIr/EIr)")] double γ,
          [ExcelArgument(Name = "Tolerance", Description = "Load increment = tol·Ncr0 [default = 0.0001]")] object _tol,
          [ExcelArgument(Name = "Mode", Description = "Mode number [default = 1]")] object _mode,
          [ExcelArgument(Name = "Maximum", Description = "Maximum [default = 2.0]")] object _max)
    {
        // increments are based on N/Ncr0
        double tol = Check<double>(_tol, 0.0001);
        double mode = Check<double>(_mode, 1);
        double max = Check<double>(_max, 2);

        return ElasticBuckling.ke_crit(N => ElasticBuckling.det_asym_connection(N, κg, κr, ξ, γ), mode, tol, max);
    }
}
