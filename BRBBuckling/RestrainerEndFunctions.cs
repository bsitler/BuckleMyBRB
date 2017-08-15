using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;
using static BRBBucklingFunctions;
using System;

namespace BRBBuckling
{
    public static class RestrainerEndFunctions
    {
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: plastic moment transfer capacity
                returns {∑M,Mc,M1,M2,M3}
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_MpRHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            double rot = RestEnd_YieldDisp2RHS(Br, tr, a, fyr, E) / Lin;

            return RestEnd_MRHS(rot, Br, tr,a , Lin, fyr, Bc, tc, fyc, E);
        }
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: rotational stiffness
                returns {K1,θy1,K2,θy2}
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_KrRHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            // elastic stiffness. eqn 40
            double rot1 = RestEnd_YieldDisp1RHS(Br, tr, a, fyr, E) / Lin;
            double Kr1 = RestEnd_MRHS(rot1, Br, tr, a, Lin, fyr, Bc, tc, fyc, E)[0] / rot1;
            // plastic stiffness approximation. eqn 45
            double rot2 = RestEnd_YieldDisp2RHS(Br, tr, a, fyr, E) / Lin;
            double Kr2 = (RestEnd_MRHS(rot2, Br, tr, a, Lin, fyr, Bc, tc, fyc, E)[0] -
                    RestEnd_MRHS(rot1, Br, tr, a, Lin, fyr, Bc, tc, fyc, E)[0])
                    / (rot2 - rot1);

            return new double[] { Kr1, rot1, Kr2, rot2 };
        }
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: rotational stiffness (approximation)
                returns {K1,θpy1,K2,θy2}
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_KrRHSApprox(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Neck Breadth")]  double Bn, [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            // Trilinear
            // elastic stiffness. eqn 40
            double Kr1 = E * Br * Pow(tr, 3) * Pow(Lin, 3) / 3 / (2 * Br * Pow(a, 3) - 3 * Pow(a, 4));
            // (psuedo) flexural yield rotation approximation. eqn 45
            double rot1 = RestEnd_YieldDisp1RHS(Br, tr, a, fyr, E) / Lin;
            double rotP1 = 0.00164 * fyr / E * Br / tr * Bn / Lin;
            // plastic stiffness approximation. eqn 45
            double Kr2 = 0.11 * fyr * Pow(Br, 3) * Pow(Lin / Bn, 3);
            // in-plane yield rotation. eqn 42
            double rot2 = RestEnd_YieldDisp2RHS(Br, tr, a, fyr, E) / Lin;

            return new double[] { Kr1, rotP1, Kr2, rot2 };
        }
        //--------------------------
        // RHS Moment-rotation plot
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: moment for a given rotation
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_MRHS(
            [ExcelArgument("Rotation")] double rot, [ExcelArgument("Restrainer breadth")] double Br,
            [ExcelArgument("Restrainer thickness")]  double tr, [ExcelArgument("Lever arm")]  double a,
            [ExcelArgument("Insert length")]  double Lin, [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            // RESTRAINER CASING MOMENT TRANSFER
            // displacement where flexural yielding starts
            double dy1 = RestEnd_YieldDisp1RHS(Br, tr, a, fyr, E);
            // displacement where catenary yielding starts
            double dy2 = RestEnd_YieldDisp2RHS(Br, tr, a, fyr, E);
            // position where flexural yielding starts
            double xy1 = Min(dy1 / rot, Lin);
            // position where catenary yielding starts
            double xy2 = Min(dy2 / rot, Lin);

            // elastic portion. eqn 23
            double M1 = E * Br * Pow(tr, 3) * Pow(xy1, 3) / (3 * (2 * Br * Pow(a, 3) - 3 * Pow(a, 4))) * rot;
            // portion yielded in casing out-of-plane flexure. eqn 24
            double M2 = E * Br * Pow(tr, 3) / (2 * Br * Pow(a, 3) - 3 * Pow(a, 4)) * dy1 * (Pow(xy2, 2) - Pow(xy1, 2)) / 2
                + (2 * dy2 * fyr * tr / Sqrt(Pow(a, 2) + Pow(dy2, 2)) - E * Br * Pow(tr, 3) * dy1 / (2 * Br * Pow(a, 3) - 3 * Pow(a, 4)))
                * (rot * (Pow(xy2, 3) - Pow(xy1, 3)) / 3 / (dy2 - dy1) - dy1 * (Pow(xy2, 2) - Pow(xy1, 2)) / 2 / (dy2 - dy1));
            // portion yielded in catenary action. eqn 25
            double M3 = fyr * tr * (Lin * Sqrt(Pow(Lin, 2) + Pow(a / rot, 2)) - xy2 * Sqrt(Pow(xy2, 2) + Pow(a / rot, 2))
                - Pow(a / rot, 2) * Log((rot * Lin + Sqrt(Pow(a, 2) + Pow(rot * Lin, 2))) / (rot * xy2 + Sqrt(Pow(a, 2) + Pow(rot * xy2, 2)))));

            // CORE MOMENT TRANSFER
            double Mc = 0;
            if (Bc > 0 && tc > 0 && fyc > 0)
                if (rot <= 6 * fyc / E)
                    Mc = E * Bc * Pow(tc, 2) / 36 * rot;
                else
                    Mc = Bc * fyc * Pow(tc, 2) / 4 - 3 * Bc * Pow(tc, 2) * Pow(fyc, 3) / Pow(E, 2) / Pow(rot, 2);

            return new double[] { M1 + M2 + M3 + Mc, Mc, M1, M2, M3 };
        }
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: flexural yield displacement
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        private static double RestEnd_YieldDisp1RHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]double a, 
            [ExcelArgument("Yield strength")]  double fy, [ExcelArgument("Youngs modulus")]  double E)
        {
            // RHS casing flexural yield. Eqn 18 (Matsui 2010)
            return fy / E / 3 / tr * (2 * Br * Pow(a, 2) - 3 * Pow(a, 3)) / (Br - a);
        }
        private static double RestEnd_YieldDisp2RHS(
            [ExcelArgument("Restrainer width")] double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")] double a, 
            [ExcelArgument("Yield strength")]  double fy, [ExcelArgument("Youngs modulus")]  double E)
        {
            // RHS casing catenary yield. Eqn 20 (Matsui 2010)
            return Br * Sqrt(Pow(fy / 2 / E, 2) + a / Br * fy / E);
        }

            [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: lever arm
            a=(Br-Bn)/4
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double RestEnd_aRHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Neck width")]  double Bn)
        {
            // Assumed offset from outer width of transverse load. Fig 6 (Matsui 2010)
            return (Br - Bn) / 4;
        }
        [ExcelFunction(Description = @"Mortar-filled RHS restrainer end: lever arm
            Balance of corner mortar strut&tie node and restrainer yield
                a=Py/(4γ·fc), Py=E·Br·tr²/3(Bra-a²)")]
        public static double RestEnd_aRHS2(
                [ExcelArgument("Restrainer width")]  double Br,
                [ExcelArgument("Restrainer thickness")]  double tr,
                [ExcelArgument("Lever arm tolerance [default = Br/100]")] object _da,
                [ExcelArgument("Node confinement factor [default = 0.5 C-C-C node]")] object _γ,
                [ExcelArgument("Mortar cylinder strength [default = 30MPa]")] object _fc,
                [ExcelArgument("Steel yield strength [default = 300MPa]")] object _fyr)
        {
            double fc = Check<double>(_fc, 30);
            double fyr = Check<double>(_fyr, 300);
            double γ = Check<double>(_γ, 0.5);
            double da = Check<double>(_da, Br/100);
            
            Func<double, double> SquashLimit = (a) => {
                var Py = fyr * Br * Pow(tr, 2) / 3 / (Br * a - Pow(a, 2));
                return Py / (2 * γ * fc) / a; };

            return BRBBuckling.NLEquationSolver.BruteForce(
                SquashLimit, 1.0, da, Br/2, da);
        }
        //--------------------------
        // CHS Plastic moment capacity and initial stiffness
        // set fyc=0 to ignore core contribution (typically yielded due to axial demands)
        [ExcelFunction(Description = @"Mortar-filled CHS restrainer end: plastic moment transfer capacity
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_MpCHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            double rot = (RestEnd_YieldDispCHS(Br, tr, a, fyr, E) - RestEnd_RefDispCHS(Br, tr, a)) / Lin;

            return RestEnd_MCHS(rot, Br, tr,a, Lin, fyr, Bc, tc, fyc, E);
        }
        [ExcelFunction(Description = @"Mortar-filled CHS restrainer end: rotational stiffness
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_KrCHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")]  double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            double dy = RestEnd_YieldDispCHS(Br, tr, a, fyr, E);
            double d0 = RestEnd_RefDispCHS(Br, tr, a);

            //Dim rot , double M As Variant, K As Double
            //rot = (dy - d0) / Lin
            //M = MCHS(rot, Br, tr, Bn, Lin, fyr, 0, 0, 0, E) // equal to eqn 38
            //K = M(1) / rot

            // elastic stiffness. eqn 38
            double K1 = fyr * tr * Pow(Lin, 3) * 2 / 3 / Sqrt(Pow(a, 2) + Pow(dy, 2)) * dy / (dy - d0);
            double K2 = (dy - d0) / Lin;

            return new double[] { K1, K2 };
        }
        //--------------------------
        // CHS Moment-rotation plot
        [ExcelFunction(Description = @"Mortar-filled CHS restrainer end: moment for a given rotation
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double[] RestEnd_MCHS(
            [ExcelArgument("Rotation")] double rot,
            [ExcelArgument("Restrainer breadth")] double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")]  double a, [ExcelArgument("Insert length")]  double Lin,
            [ExcelArgument("Yield strength")]  double fyr,
            [ExcelArgument("Core breadth")] double Bc, [ExcelArgument("Core thickness")]  double tc,
            [ExcelArgument("Yield strength")]  double fyc, [ExcelArgument("Youngs modulus")]  double E)
        {
            // RESTRAINER CASING MOMENT TRANSFER
            // position where flexural yielding starts
            double dy = RestEnd_YieldDispCHS(Br, tr, a, fyr, E);
            double d0 = RestEnd_RefDispCHS(Br, tr, a);
            double D = Max(rot * Lin + d0, dy);
            double xy = Min((dy - d0) / rot, Lin);

            // elastic portion. eqn 33, 36
            double M1 = 2 * fyr * tr * dy / Sqrt(Pow(a, 2) + Pow(dy, 2)) / 3 / (dy - d0) * rot * Pow(xy, 3);
            // portion yielded in casing out-of-plane flexure. eqn 37
            double M2 = 2 * fyr * tr / Pow(rot, 2) *
                ((D / 2 - d0) * Sqrt(Pow(a, 2) + Pow(D, 2))
                - (dy / 2 - d0) * Sqrt(Pow(a, 2) + Pow(dy, 2))
                - Pow(a, 2) / 2 * Log((D + Sqrt(Pow(a, 2) + Pow(D, 2))) / (dy + Sqrt(Pow(a, 2) + Pow(dy, 2)))));

            // CORE MOMENT TRANSFER
            double Mc = 0;
            if (Bc > 0 && tc > 0 && fyc > 0)
                if (rot <= 6 * fyc / E)
                    Mc = E * Bc * Pow(tc, 2) / 36 * rot;
                else
                    Mc = Bc * fyc * Pow(tc, 2) / 4 - 3 * Bc * Pow(tc, 2) * Pow(fyc, 3) / Pow(E, 2) / Pow(rot, 2);

            return new double[] { M1 + M2 + Mc, Mc, M1, M2 };
        }

        private static double RestEnd_RefDispCHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")] double a)
        {
            return Sqrt(a * (Br - a) - tr * (Br - tr));
        }
        private static double RestEnd_YieldDispCHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Restrainer thickness")]  double tr,
            [ExcelArgument("Lever arm")] double a,
            [ExcelArgument("Yield strength")]  double fyr, [ExcelArgument("Youngs modulus")]  double E)
        {
            // RHS casing flexural yield. Eqn 32 (Matsui 2010)
            return Sqrt(Pow(PI * Br * fyr / 4 / E + Br / 2 * Acos((Br - 2 * a) / (Br - 2 * tr)), 2) - Pow(a, 2));
        }
        [ExcelFunction(Description = @"Mortar-filled CHS restrainer end: lever arm
            a=(Br-Bn)/2
            Matsui 2010: Effective buckling length of BRB considering rotational stiffness at restrainer ends. 7CUEE@Tokyo Tech")]
        public static double RestEnd_aCHS(
            [ExcelArgument("Restrainer width")]  double Br, [ExcelArgument("Neck width")]  double Bn)
        {
            // Assumed offset from outer width of transverse load. Fig 10 (Matsui 2010)
            return (Br - Bn) / 2;
        }
    }
}