using static System.Math;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;
using static BRBBuckling.OptionalParams;

public static class SectionPropertiespublic
{
    #region Area
    [ExcelFunction(Description = @"Section gross area")]
    public static double area(string Section,
        [ExcelArgument("Depth", Name = "D")] object _D,
        [ExcelArgument("Breadth", Name = "B")] object _B,
        [ExcelArgument("Flange/horiz plate thickness", Name = "tf_h")] object _tf_h,
        [ExcelArgument("Web/vert plate thickness", Name = "tw_v")] object _tw_v)
    {
        double D = Check<double>(_D, 0);
        double t = Check<double>(_tf_h, 0);
        double B = Check<double>(_B, 0);
        double tw = Check<double>(_tw_v, 0);
        double tf = t; double th = t; double tv = tw;

        switch (Section)
        {
            case "I":
            case "H":
            case "ISection":
            case "HSection":
                return Area_ISection(D, B, tf, tw);
            case "RHS":
                return Area_RHS(D, B, tf, tw);
            case "SHS":
                return Area_SHS(D, t);
            case "CHS":
                return Area_CHS(D, t);
            case "Rectangle":
                return Area_Rectangle(D, B);
            case "T":
            case "TSection":
                return Area_TSection(D, B, tf, tw);
            case "Angle":
                return Area_Angle(D, B, tf, tw);
            case "Cruciform":
                return Area_Cruciform(D, B, tv, th);
            case "Channel":
                return Area_Channel(D, B, tf, tw);
            default: return 0;
        }
    }
    [ExcelFunction(Description = @"Circular hollow section gross area
                                    A=π(D²-(D-2t)²)/4")]
    public static double Area_CHS([ExcelArgument("Diameter")] double D, [ExcelArgument("Thickness")] double t)
    {
        return PI * (Pow(D, 2) - Pow(D - 2 * t, 2)) / 4;
    }
    [ExcelFunction(Description = @"Rectangular hollow section gross area
                                    A=D·B-(D-2tf)(B-2tw)")]
    public static double Area_RHS([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Top/Btm plate thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return D * B - (D - 2 * tf) * (B - 2 * tw);
    }
    [ExcelFunction(Description = @"Square hollow section gross area
                                    A=D·B-(D-2t)²")]
    public static double Area_SHS([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return Area_RHS(D, D, t, t);
    }
    [ExcelFunction(Description = @"Rectangle section gross area
                                    A=D·B")]
    public static double Area_Rectangle([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B)
    {
        return D * B;
    }
    [ExcelFunction(Description = @"I/H section gross area
                                    A=2B·tf+(D-2tf)tw")]
    public static double Area_ISection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return 2 * B * tf + (D - 2 * tf) * tw;
    }
    [ExcelFunction(Description = @"Channel section gross area
                                    A=2B·tf+(D-2tf)tw")]
    public static double Area_Channel([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return Area_ISection(D, B, tf, tw);
    }
    [ExcelFunction(Description = @"Cruciform section gross area
                                    A=D·tv+(B-tv)th")]
    public static double Area_Cruciform([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Vert. plate thickness")] double tv, [ExcelArgument("Horiz. plate thickness")] double th)
    {
        return D * tv + (B - tv) * th;
    }
    [ExcelFunction(Description = @"T section gross area
                                    A=(D-tf)·tw+B·tf")]
    public static double Area_TSection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return (D - tf) * tw + B * tf;
    }
    [ExcelFunction(Description = @"Angle gross area
                                    A=(D-tf)·tw+B·tf")]
    public static double Area_Angle([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange (horiz plate) thickness")] double tf, [ExcelArgument("Web (vert plate) thickness")] double tw)
    {
        return (D - tf) * tw + B * tf;
    }
    #endregion

    #region Iyy
    [ExcelFunction(Description = @"Section (major axis) moment of inertia")]
    public static double Iyy(string Section,
        [ExcelArgument("Depth", Name = "D")] object _D,
        [ExcelArgument("Breadth", Name = "B")] object _B,
        [ExcelArgument("Flange/horiz plate thickness", Name = "tf_h")] object _tf_h,
        [ExcelArgument("Web/vert plate thickness", Name = "tw_v")] object _tw_v)
    {
        double D = Check<double>(_D, 0);
        double t = Check<double>(_tf_h, 0);
        double B = Check<double>(_B, 0);
        double tw = Check<double>(_tw_v, 0);
        double tf = t; double th = t; double tv = tw;

        switch (Section)
        {
            case "I":
            case "H":
            case "ISection":
            case "HSection":
                return Iyy_ISection(D, B, tf, tw);
            case "RHS":
                return Iyy_RHS(D, B, tf, tw);
            case "SHS":
                return Iyy_SHS(D, t);
            case "CHS":
                return Iyy_CHS(D, t);
            case "Rectangle":
                return Iyy_Rectangle(D, B);
            case "T":
            case "TSection":
                return Iyy_TSection(D, B, tf, tw);
            case "Angle":
                return Iyy_Angle(D, B, tf, tw);
            case "Channel":
                return Iyy_Channel(D, B, tf, tw);
            case "Cruciform":
                return Iyy_Cruciform(D, B, tv, th);
            default: return 0;
        }
    }
    [ExcelFunction(Description = @"Circular hollow section moment of inertia
                                    Iyy=π(D⁴-(D-2t)⁴)/64")]
    public static double Iyy_CHS([ExcelArgument("Diameter")] double D, [ExcelArgument("Thickness")] double t)
    {
        return PI * (Pow(D, 4) - Pow(D - 2 * t, 4)) / 64;
    }
    [ExcelFunction(Description = @"Rectangular hollow section (major axis) moment of inertia
                                    Iyy=(B·D³-(B-2tw)(D-2tf)³)/12")]
    public static double Iyy_RHS([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Top/Btm plate thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return (B * Pow(D, 3) - (B - 2 * tw) * Pow(D - 2 * tf, 3)) / 12;
    }
    [ExcelFunction(Description = @"Square hollow section moment of inertia
                                    Iyy=(D⁴-(D-2t)⁴)/12")]
    public static double Iyy_SHS([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return Iyy_RHS(D, D, t, t);
    }
    [ExcelFunction(Description = @"Rectangle section (major axis) moment of inertia
                                    Iyy=t·D³/12")]
    public static double Iyy_Rectangle([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return t * Pow(D, 3) / 12;
    }
    [ExcelFunction(Description = @"I/H section (major axis) moment of inertia
                                    Iyy=(B·D³-(B-tw)(D-2tf)³)/12")]
    public static double Iyy_ISection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return (B * Pow(D, 3) - (B - tw) * Pow(D - 2 * tf, 3)) / 12;
    }
    [ExcelFunction(Description = @"Channel section (major axis) moment of inertia
                                    Iyy=(B·D³-(B-tw)(D-2tf)³)/12")]
    public static double Iyy_Channel([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return Iyy_ISection(D, B, tf, tw);
    }
    [ExcelFunction(Description = @"Cruciform section (major axis) moment of inertia
                                    Iyy=(td·D³+(B-td)tb³)/12")]
    public static double Iyy_Cruciform([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Top/Btm thickness")] double tv, [ExcelArgument("Side thickness")] double th)
    {
        return (tv * Pow(D, 3) + (B - tv) * Pow(th, 3)) / 12;
    }
    [ExcelFunction(Description = @"T section (major axis) moment of inertia
                                    Iyy=(Btd·D³+(B-td)tb³)/12
                                                +(B·th³+tv(D-th)³)/12+Bth·(D-yn-th/2)²
                                                +tv(D-th)(yn-(D-th)/2)²")]
    public static double Iyy_TSection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        double yn = yn_TSection(D, B, tf, tw);
        return (B * Pow(tf, 3) + tw * Pow(D - tf, 3)) / 12
            + B * tf * Pow(yn - tf / 2, 2)
            + tw * (D - tf) * Pow(yn - (D + tf) / 2, 2);
    }
    [ExcelFunction(Description = @"Angle section (major axis) moment of inertia
                                    Iyy=(refer Iyy_TSection)")]
    public static double Iyy_Angle([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return Iyy_TSection(D, B, tf, tw);
    }
    #endregion

    #region Depth to neutral axis yn
    [ExcelFunction(Description = @"Section (major axis) moment of inertia")]
    public static double yn(string Section,
        [ExcelArgument("Depth", Name = "D")] object _D,
        [ExcelArgument("Breadth", Name = "B")] object _B,
        [ExcelArgument("Flange/horiz plate thickness", Name = "tf_h")] object _tf_h,
        [ExcelArgument("Web/vert plate thickness", Name = "tw_v")] object _tw_v)
    {
        double D = Check<double>(_D, 0);
        double t = Check<double>(_tf_h, 0);
        double B = Check<double>(_B, 0);
        double tw = Check<double>(_tw_v, 0);
        double tf = t; double th = t; double tv = tw;

        switch (Section)
        {
            case "I":
            case "H":
            case "ISection":
            case "HSection":
            case "Cruciform":
            case "RHS":
            case "SHS":
            case "Rectangle":
                return D / 2;
            case "T":
            case "TSection":
                return yn_TSection(D, B, tf, tw);
            case "Angle":
                return yn_Angle(D, B, tf, tw);
            default: return 0;
        }
    }
    [ExcelFunction(Description = @"T section depth to neutral axis
                                                yn=(Bth(D-th/2)+tv(D-th)²/2)/Area")]
    public static double yn_TSection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        double A = Area_TSection(D, B, tf, tw);
        return (B * Pow(tf, 2) + tw * Pow(D - tf, 2) + 2 * tf * tw * (D - tf)) / 2 / A;
    }
    [ExcelFunction(Description = @"Angle depth to neutral axis
                                                yn=Bth(D-th/2)+tv(D-th)²/2)/Area")]
    public static double yn_Angle([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw) {
        return yn_TSection(D, B, tf, tw);
    }
    #endregion

    #region Composite Iyy
    [ExcelFunction(Description = @"Parallel axis theorem
                                            I=I+A·yn²")]
    public static double ParallelAxisTheorem(
        [ExcelArgument("Component second moment of areas")] double[] I,
        [ExcelArgument("Component areas")] double[] A,
        [ExcelArgument("Component neutral axis to datum")] double[] yn)
    {
        double PAT = I.Sum();
        if (A.Length != yn.Length) return 0;
        for (int i = 0; i < A.Length; i++)
            PAT += A[i] * Pow(yn[i], 2);
        return PAT;
    }
    private static double ParallelAxisTheorem(double I,double A,double yn)
    {
        return I + A * Pow(yn, 2);
    }
    //[ExcelFunction(Description = @"Double T section (major axis) moment of inertia")]
    //public static double Iyy_DoubleTSection(
    //    [ExcelArgument("T Depth")] double D, 
    //    [ExcelArgument("T Vert. plate thickness")] double tv, 
    //    [ExcelArgument("T Breadth")] double B, 
    //    [ExcelArgument("T Horiz. plate thickness")] double th,
    //    [ExcelArgument("Separation gap")] double gap,
    //    [ExcelArgument("Do T sections act compositely [default=true]",Name ="isComposite")] object _isComposite)
    //{
    //    bool isComposite = Check<bool>(_isComposite, true);
    //    double Iyy = Iyy_TSection(D, tv, B, th);
    //    if (!isComposite) return 2 * Iyy;

    //    double A = Area_TSection(D, tv, B, th);
    //    double yn = yn_TSection(D, tv, B, th)+gap/2;
    //    return ParallelAxisTheorem(new double[] { Iyy, Iyy }, new double[] { A, A }, new double[] { yn, yn });
    //}
    //[ExcelFunction(Description = @"Double rectangle (major axis) moment of inertia")]
    //public static double Iyy_DoubleRectangle(
    //    [ExcelArgument("Depth")] double D,
    //    [ExcelArgument("Breadth")] double B,
    //    [ExcelArgument("Separation gap")] double gap,
    //    [ExcelArgument("Do T sections act compositely [default=true]", Name = "isComposite")] object _isComposite)
    //{
    //    bool isComposite = Check<bool>(_isComposite, true);
    //    double Iyy = Iyy_Rectangle(D, B);
    //    if (!isComposite) return 2 * Iyy;

    //    double A = Area_Rectangle(D, B);
    //    double yn = D/2 + gap / 2;
    //    return 2*ParallelAxisTheorem(Iyy ,A ,yn );
    //}
    //[ExcelFunction(Description = @"Double RHS (major axis) moment of inertia")]
    //public static double Iyy_DoubleRHS(
    //    [ExcelArgument("Depth")] double D,
    //    [ExcelArgument("Breadth")] double B,
    //    [ExcelArgument("Thickness")] double t,
    //    [ExcelArgument("Separation gap")] double gap,
    //    [ExcelArgument("Do T sections act compositely [default=true]", Name = "isComposite")] object _isComposite)
    //{
    //    bool isComposite = Check<bool>(_isComposite, true);
    //    double Iyy = Iyy_RHS(D, t, B, t);
    //    if (!isComposite) return 4 * Iyy;

    //    double A = Area_RHS(D, t, B, t);
    //    double yn = D / 2 + gap / 2;
    //    return 2 * ParallelAxisTheorem(Iyy, A, yn);
    //}
    //[ExcelFunction(Description = @"Quad RHS (major axis) moment of inertia")]
    //public static double Iyy_QuadRHS (
    //    [ExcelArgument("Depth")] double D,
    //    [ExcelArgument("Breadth")] double B,
    //    [ExcelArgument("Thickness")] double t,
    //    [ExcelArgument("Separation gap")] double gap,
    //    [ExcelArgument("Do T sections act compositely [default=true]", Name = "isComposite")] object _isComposite)
    //{
    //    bool isComposite = Check<bool>(_isComposite, true);
    //    double Iyy = Iyy_RHS(D,t,B,t);
    //    if (!isComposite) return 4 * Iyy;

    //    double A = Area_RHS(D,t,B,t);
    //    double yn = D / 2 + gap / 2;
    //    return 4*ParallelAxisTheorem(Iyy ,A, yn );
    //}
 //   [ExcelFunction(Description = @"Quad angle (major axis) moment of inertia")]
 //   public static double Iyy_QuadAngle(
 //       [ExcelArgument("Depth")] double D,
 //       [ExcelArgument("Breadth")] double B,
 //       [ExcelArgument("Thickness")] double t,
 //       [ExcelArgument("Separation gap")] double gap,
 //       [ExcelArgument("Do T sections act compositely [default=true]", Name = "isComposite")] object _isComposite)
 //   {
 //       return Iyy_DoubleTSection(D, 2 * t, 2 * B, t, gap, _isComposite);
 //   }
    #endregion

    #region Izz
    [ExcelFunction(Description = @"Section (minor axis) moment of inertia")]
    public static double Izz(string Section,
         [ExcelArgument("Depth/Diameter", Name = "D")] object _D,
        [ExcelArgument("Breadth", Name = "B")] object _B,
        [ExcelArgument("Flange/horiz plate thickness", Name = "tf_h")] object _tf_h,
        [ExcelArgument("Web/vert plate thickness", Name = "tw_v")] object _tw_v)
    {
        double D = Check<double>(_D, 0);
        double t = Check<double>(_tf_h, 0);
        double B = Check<double>(_B, 0);
        double tw = Check<double>(_tw_v, 0);
        double tf = t; double th = t; double tv = tw;


        switch (Section)
        {
            case "I":
            case "H":
            case "ISection":
            case "HSection":
                return Izz_ISection(D, B, tf, tw);
            case "Cruciform":
                return Izz_Cruciform(D, B, tv, th);
            case "RHS":
                return Izz_RHS(D, B, tf, tw);
            case "SHS":
                return Izz_SHS(D, t);
            case "CHS":
                return Izz_CHS(D, t);
            case "Rectangle":
                     return Izz_Rectangle(D, B);
            default: return 0;
        }
    }
    [ExcelFunction(Description = @"Circular hollow section moment of inertia
                                    Izz=π(D⁴-(D-2t)⁴)/64")]
    public static double Izz_CHS([ExcelArgument("Diameter")] double D, [ExcelArgument("Thickness")] double t)
    {
        return Iyy_CHS(D, t);
    }
    [ExcelFunction(Description = @"Rectangular hollow section (minor axis) moment of inertia
                                    Izz=(D·B³-(D-2tf)(B-2tw)³)/12")]
    public static double Izz_RHS([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Top/Btm plate thickness")] double tf,[ExcelArgument("Web thickness")] double tw)
    {
        return Iyy_RHS(B, D, tw,  tf);
    }
    [ExcelFunction(Description = @"Square hollow section moment of inertia
                                    Izz=(D⁴-(D-2t)⁴)/12")]
    public static double Izz_SHS([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return Iyy_SHS(D, t);
    }
    [ExcelFunction(Description = @"Rectangle section (minor axis) moment of inertia
                                    Izz=D·t³/12")]
    public static double Izz_Rectangle([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return Iyy_Rectangle(t, D);
    }
    [ExcelFunction(Description = @"I/H section (minor axis) moment of inertia
                                    Iyy=(2tf·B³+(D-2tf)tw³)/12")]
    public static double Izz_ISection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        return Iyy_Cruciform(B, D, 2 * tf,  tw);
    }
    [ExcelFunction(Description = @"Channel section (minor axis) moment of inertia
                                    Roarde Tbl A.1(5) Izz=D·B³/3-(B-tw)³(D-2tf)/3-A(B-c)²
                                    where A=2B·tf+(D-2tf)tw,
                                    c=(D·tw²+2tf(B²-tw²))/2(tw·D+2tf(B-tw))")]
    public static double Izz_Channel([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf,  [ExcelArgument("Web thickness")] double tw)
    {
        // Roarka Table A.1 #5
        double yc = (D * Pow(tw, 2) + 2 * tf * (B - tw) * (tw + B)) / (2 * (tw * D + 2 * tf * (B - tw)));
        double A = Area_Channel(D, B, tf,  tw);
        return D / 3 * Pow(B, 3) - Pow(B - tw, 3) / 3 * (D - 2 * tf) - A * Pow(B - yc, 2);
    }
    [ExcelFunction(Description = @"Cruciform section (minor axis) moment of inertia
                                    Izz=(tb·B³+(D-tb)td³)/12")]
    public static double Izz_Cruciform([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Vert. plate thickness")] double tv,[ExcelArgument("Horiz. plate thickness")] double th)
    {
        return Iyy_Cruciform(B, D, th, tv);
    }
    #endregion

    #region Radius of gyration
    [ExcelFunction(Description = @"Rectangle radius of gyration (major axis)
                                    ryy=D/√12")]
    public static double ryy_Rectangle([ExcelArgument("Depth")] double D)
    {
        return D/Sqrt(12);
    }
    #endregion

    #region J
    [ExcelFunction(Description = @"Torsional moment of inertia reduced for warping")]
    public static double J(string Section,
        [ExcelArgument("Depth", Name = "D")] object _D,
        [ExcelArgument("Breadth", Name = "B")] object _B,
        [ExcelArgument("Flange/horiz plate thickness", Name = "tf_h")] object _tf_h,
        [ExcelArgument("Web/vert plate thickness", Name = "tw_v")] object _tw_v)
    {
        double D = Check<double>(_D, 0);
        double t = Check<double>(_tf_h, 0);
        double B = Check<double>(_B, 0);
        double tw = Check<double>(_tw_v, 0);
        double tf = t; double th = t; double tv = tw;

        switch (Section)
        {
            case "I":
            case "H":
            case "ISection":
            case "HSection":
                return J_ISection(D, B, tf, tw);
            case "RHS":
                return J_RHS(D, B, tv,  th);
            case "SHS":
                return J_SHS(D, t);
            case "CHS":
                return J_CHS(D, t);
            default: return 0;
        }
    }

    [ExcelFunction(Description = @"Circular hollow section torsional moment of inertia reduced for warping
                                    Roarke Tbl 10.1.13 with U=πD, J=πt(D-t)³′⁴
                                    J≈π·t(D-t)³/4")]
    public static double J_CHS([ExcelArgument("Diameter")] double D, [ExcelArgument("Thickness")] double t)
    {
        // Roarke tbl 10.1.13, with U = pi*D,  J = pi*t(D-t)^3/4U = pi*t(D-t)^3/4
        return PI * t * Pow(D - t, 3) / 4;
    }
    [ExcelFunction(Description = @"Rectangular hollow section torsional moment of inertia reduced for warping
                                    Roarke Tbl 10.1.16
                                    J≈2tf·tw(B-tw)²(D-tf)²/(B·tw+D·tf-tw²-tf²)")]
    public static double J_RHS([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Top/Btm plate thickness")] double tf,  [ExcelArgument("Side plate thickness")] double tw)
    {
        // Roarke tbl 10.1.16
        return (2 * tf * tw * Pow(B - tw, 2) * Pow(D - tf, 2)) / (B * tw + D * tf - Pow(tw, 2) - Pow(tf, 2));
    }
    [ExcelFunction(Description = @"Square hollow section torsional moment of inertia reduced for warping
                                    Roarke Tbl 10.1.16
                                    J≈2t²(B-tb)⁴/(2D·t-2t²)")]
    public static double J_SHS([ExcelArgument("Depth")] double D, [ExcelArgument("Thickness")] double t)
    {
        return J_RHS(D, D, t,  t);
    }
    [ExcelFunction(Description = @"I/H Section torsional moment of inertia reduced for warping
                                    Roarke Tbl 10.2.6
                                    J≈(2B·tf³+D·tw³)/3")]
    public static double J_ISection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf,  [ExcelArgument("Web thickness")] double tw)
    {
        // Roarke Tbl 10.2.6
        return (2 * B * Pow(tf, 3) + D * Pow(tw, 3)) / 3;
    }
    [ExcelFunction(Description = @"I/H Section warping constant
                                    Roarke Tbl 10.2.6
                                    Cw≈0.5·B·tf³(D-tf)²/12")]
    public static double Cw_ISection([ExcelArgument("Depth")] double D, [ExcelArgument("Breadth")] double B, [ExcelArgument("Flange thickness")] double tf, [ExcelArgument("Web thickness")] double tw)
    {
        // Roarke Tbl 10.2.6
        return 0.5 * (tf*Pow( B, 3) / 12) * Pow(D - tf, 2);
    }
    #endregion
}