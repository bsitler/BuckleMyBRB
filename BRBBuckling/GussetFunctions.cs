using static System.Math;
using ExcelDna.Integration;
using static BRBBuckling.OptionalParams;
namespace BRBBuckling
{
    public static class GussetFunctions
    {

        #region Kinoshita's Method - Gusset Type A
        [ExcelFunction(Description = @"Gusset effective yield line length (Gusset Type A)

            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Kg_TypeA(
                [ExcelArgument("Aspect ratio")] double AR,
                [ExcelArgument("Angle [rad] between yield line normal and BRB", Name = "θ")] double theta,
                [ExcelArgument("Gusset thickness")] double tg,
                [ExcelArgument("Youngs modulus [default = 205000MPa]", Name = "E")] object _E,
                [ExcelArgument("Poisson ratio (elastic) [default = 0.3]", Name = "ν")] object _v)
        {
            double E = Check<double>(_E, 205000);
            double v = Check<double>(_v, 0.3);

            double D = E * (Pow(tg, 3)) / 12 / (1 - Pow(v, 2));

            return 2 * D * AR / Cos(theta);
        }

        [ExcelFunction(Description = @"Gusset effective yield line length
            where center rib does not penetrate the yield line (Gusset Type A)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Ly_TypeA(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            var gusset = new KinoshitaGusset(angle, Wt1, Wt2, Ls, Lg, Lgx, 0, 0, φ1, φ2);
            return gusset.p2.DistanceTo(gusset.p4);
        }
        #endregion

        #region Kinoshita's Method - Aspect Ratio (Gusset Type A)
        // GUSSET TYPE A (CENTER STIFFENER DOES NOT CROSS YIELD LINE)
        [ExcelFunction(Description = @"Overall aspect ratio
            AR=L01/(acHa+acHc)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_AR_TypeA(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            var gusset = new KinoshitaGusset(angle, Wt1, Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2);
            return gusset.AspectRatio(gusset.p2, gusset.p1, gusset.p4, gusset.p0);
        }
        [ExcelFunction(Description = @"Overall internal angle (Gusset Type A)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_theta_TypeA(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            var gusset = new KinoshitaGusset(angle, Wt1, Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2);
            return PI / 2 - angle - gusset.p4.Angle(gusset.p2);
        }
        #endregion

        #region Kinoshita's Method - Gusset Type B or C
        // Gusset rotational stiffness
        // refer Kinoshita (2008) Eqns 7-9c
        [ExcelFunction(Description = @"Gusset effective yield line length 
            where center rib penetrates yield line (Gusset Types B&C)
            Ly = (Lb+L13·(|cotθA|+|cotθB|+|cscθA·cosφA·sec(θA-φA)|+|cscθB·cosφB·sec(θB-φB)|))
            if φA=φB=0, Mgy = (Lb+L13·(|cotθA|+|cotθB|+|2csc2θA|+|2csc2θB|))
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Ly_TypeBC(
             [ExcelArgument("Internal angle (panel A)", Name = "θA")] double thetaA,
             [ExcelArgument("Internal angle (panel B)", Name = "θB")] double thetaB,
             [ExcelArgument("Length of center rib")] double L13,
              [ExcelArgument("Bolt width [default=0]", Name = "Lb")] object _Lb,
             [ExcelArgument("Effective taper angle (panel A) [default=0 rad]", Name = "φA'")] object _phiA,
             [ExcelArgument("Effective taper angle (panel B) [default=0 rad]", Name = "φB'")] object _phiB)
        {
            double Lb = Check<double>(_Lb, 0);
            double phiA = Check<double>(_phiA, 0);
            double phiB = Check<double>(_phiB, 0);

            return (Lb + L13 * (Abs(1 / Tan(thetaA)) + Abs(1 / Tan(thetaB))
                + Abs(Cos(phiA) / Sin(thetaA) / Cos(thetaA - phiA)) + Abs(Cos(phiB) / Sin(thetaB) / Cos(thetaB - phiB))));
        }


        [ExcelFunction(Description = @"Gusset yield moment (no reduction for NM interaction)
            Mgy = fy·tg²/6·Ly
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Mgy(
              [ExcelArgument("Effective yield line length")] double Ly,
          [ExcelArgument("Gusset thickness")] double tg,
              [ExcelArgument("Yield strength")]  double fy)
        {
            double Z = Ly * Pow(tg, 2) / 4;
            return fy * Z;
        }
        [ExcelFunction(Description = @"Gusset edge stiffener yield moment (no reduction for NM interaction)
            Msy = fy·ts·D²/6·secθ
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Msy(
          [ExcelArgument("Angle between stiffener and BRB", Name = "θ")] double theta,
          [ExcelArgument("Edge stiffener depth")] double Ds,
          [ExcelArgument("Edge stiffener thickness")] double ts,
          [ExcelArgument("Yield strength")]  double fy)
        {
            double Z = ts * Pow(Ds, 2) / 4;
            return fy * Z / Cos(theta);
        }

        // Gusset rotational stiffness
        // refer Kinoshita (2008) Eqns 7-9c
        [ExcelFunction(Description = @"Gusset rotational stiffness
            Kg=2D·(ARab(|cotθA|+|cotθB|)²+ARac·csc²θA+ARbd·csc²θB)
            where D=E(tg³+n·ts³)/12(1-ν²) (Kirchhoff plate stiffness)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Kg(
                [ExcelArgument("Panel AB Aspect ratio (L/(H1+H2)")] double ARab,
                [ExcelArgument("Panel AC Aspect ratio (L/(H1+H2)")] double ARac,
                [ExcelArgument("Panel BD Aspect ratio (L/(H1+H2)")] double ARbd,
                [ExcelArgument("Internal angle (panel A)")] double thetaA,
                [ExcelArgument("Internal angle (panel B)")] double thetaB,
                [ExcelArgument("Gusset thickness")] double tg,
                [ExcelArgument("Youngs modulus [default = 205000MPa]", Name = "E")] object _E,
                [ExcelArgument("Poisson ratio (elastic) [default = 0.3]", Name = "ν")] object _v)
        {
            double E = Check<double>(_E, 205000);
            double v = Check<double>(_v, 0.3);

            double D = E * (Pow(tg, 3)) / 12 / (1 - Pow(v, 2));

            return 2 * D * (ARab * Pow(Abs(1 / Tan(thetaA)) + Abs(1 / Tan(thetaB)), 2)
                + ARac * Pow(1 / Sin(thetaA), 2) + ARbd * Pow(1 / Sin(thetaB), 2));
        }
        [ExcelFunction(Description = @"Gusset rotational stiffness: panel AB yield line
            K=2AR·D·(|cotθA|+|cotθB|)²
            where D=E(tg³+n·ts³)/12(1-ν²) (Kirchhoff plate stiffness)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Kab(
                [ExcelArgument("Aspect ratio (L/(H1+H2)")] double AR,
                [ExcelArgument("Internal angle (panel A)")] double thetaA,
                [ExcelArgument("Internal angle (panel B)")] double thetaB,
                [ExcelArgument("Gusset thickness")] double tg,
                [ExcelArgument("Number of splice plates [default = 0]")] object _n,
                [ExcelArgument("Splice thickness [default = 0]")] object _ts,
                [ExcelArgument("Youngs modulus [default = 205000MPa]", Name = "E")] object _E,
                [ExcelArgument("Poisson ratio (elastic) [default = 0.3]", Name = "ν")] object _v)
        {
            double n = Check<double>(_n, 0);
            double ts = Check<double>(_ts, 0);
            double E = Check<double>(_E, 205000);
            double v = Check<double>(_v, 0.3);

            double D = E * (Pow(tg, 3) + n * Pow(ts, 3)) / 12 / (1 - Pow(v, 2));

            if (thetaA == 0 || thetaB == 0)
                return 2 * AR * D;
            else
                return 2 * AR * Pow(Abs(1 / Tan(thetaA)) + Abs(1 / Tan(thetaB)), 2) * D;
        }
        [ExcelFunction(Description = @"Gusset rotational stiffness: panel AC yield line
            K=2AR·D·csc²θA
            where D=E(tg³+n·ts³)/12(1-ν²) (Kirchhoff plate stiffness)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Kac(
             [ExcelArgument("Aspect ratio (L/(H1+H2)")] double AR,
                 [ExcelArgument("Internal angle (panel A)")] double thetaA,
             [ExcelArgument("Gusset thickness")] double tg,
                 [ExcelArgument("Number of splice plates [default = 0]")] object _n,
                 [ExcelArgument("Splice thickness [default = 0]")] object _ts,
                 [ExcelArgument("Youngs modulus [default = 205000MPa]", Name = "E")] object _E,
                 [ExcelArgument("Poisson ratio (elastic) [default = 0.3]", Name = "ν")] object _v)
        {
            double n = Check<double>(_n, 0);
            double ts = Check<double>(_ts, 0);
            double E = Check<double>(_E, 205000);
            double v = Check<double>(_v, 0.3);

            double D = E * (Pow(tg, 3) + n * Pow(ts, 3)) / 12 / (1 - Pow(v, 2));

            if (thetaA == 0)
                return 2 * AR * D;
            else
                return 2 * AR * Pow(1 / Sin(thetaA), 2) * D;
        }
        [ExcelFunction(Description = @"Gusset rotational stiffness: panel BD yield line
            K=2AR·D·csc²θB
            where D=E(tg³+n·ts³)/12(1-ν²) (Kirchhoff plate stiffness)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_Kbd(
             [ExcelArgument("Aspect ratio (L/(H1+H2)")] double AR,
                 [ExcelArgument("Internal angle (panel B)")] double thetaB,
             [ExcelArgument("Gusset thickness")] double tg,
                 [ExcelArgument("Number of splice plates [default = 0]")] object _n,
                 [ExcelArgument("Splice thickness [default = 0]")] object _ts,
                 [ExcelArgument("Youngs modulus [default = 205000MPa]", Name = "E")] object _E,
                 [ExcelArgument("Poisson ratio (elastic) [default = 0.3]", Name = "ν")] object _v)
        {
            double n = Check<double>(_n, 0);
            double ts = Check<double>(_ts, 0);
            double E = Check<double>(_E, 205000);
            double v = Check<double>(_v, 0.3);

            double D = E * (Pow(tg, 3) + n * Pow(ts, 3)) / 12 / (1 - Pow(v, 2));

            if (thetaB == 0)
                return 2 * AR * D;
            else
                return 2 * AR * Pow(1 / Sin(thetaB), 2) * D;
        }
        #endregion

        #region Kinoshita's Method - Aspect Ratio (Gusset Types B or C)
        // Gusset panelization from centerline dimensions
        [ExcelFunction(Description = @"Panel AB aspect ratio
            AR=L13/(abHa+abHb)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARab(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).ARab();
        }
        [ExcelFunction(Description = @"Panel AC aspect ratio
            AR=L23/(acHa+acHc)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARac(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).ARac();
        }
        [ExcelFunction(Description = @"Panel BD aspect ratio
            AR=L34/(bdHb+bdHd)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARbd(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).ARbd();
        }

        [ExcelFunction(Description = @"Panel A internal angle
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaA(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).ThetaA();
        }
        [ExcelFunction(Description = @"Panel B internal angle
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaB(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).ThetaB();
        }
        [ExcelFunction(Description = @"Panel A tip taper angle
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_phiA(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).PhiA();
        }
        [ExcelFunction(Description = @"Panel B tip taper angle
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_phiB(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Gusset tip halfwidth to horiz edge")] double Wt1,
                [ExcelArgument("Gusset tip halfwidth to vert edge")] double Wt2,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Length of center rib")] double Lg,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx,
                [ExcelArgument("Horizontal edge stiffener [default = 0.0]")] object _percStiffener1,
                [ExcelArgument("Vertical edge stiffener [default = 0.0]")] object _percStiffener2,
                [ExcelArgument("Gusset horizontal edge taper [default = 0 rad]")] object _φ1,
                [ExcelArgument("Gusset vertical edge taper [default = 0 rad]")] object _φ2)
        {
            double Lgx = Check<double>(_Lgx, 0);
            double percStiffener1 = Check<double>(_percStiffener1, 0);
            double percStiffener2 = Check<double>(_percStiffener2, 0);
            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);

            return new KinoshitaGusset(angle, Wt1,Wt2, Ls, Lg, Lgx, percStiffener1, percStiffener2, φ1, φ2).PhiB();
        }
        #endregion
        
        #region Kinoshita's Method - Aspect Ratio Alternatives (Gusset Types B or C)
        // Gusset panelization from orthogonal gusset dimensions
        [ExcelFunction(Description = @"Panel AB aspect ratio (assuming orthogonal gusset)
            AR=L13/(abHa+abHb)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARab2(
                    [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
               [ExcelArgument("Length of center rib")] double Lf,
                    [ExcelArgument("Gusset height at column face")] double hc,
               [ExcelArgument("Gusset height at free edge (<hc)")] double he,
               [ExcelArgument("Stiffener height (<he)")] double hs,
                    [ExcelArgument("Gusset width at beam face")] double wb,
                    [ExcelArgument("Gusset width at free edge (<wb)")] double we,
               [ExcelArgument("Stiffener width (<we)")] double ws)
        {
            double Ha = abHa(angle, Lf, hc, he, hs, wb, we, ws);
            double Hb = abHb(angle, Lf, hc, he, hs, wb, we, ws);

            return Lf / (Ha + Hb);
        }
        [ExcelFunction(Description = @"Panel AC aspect ratio (assuming orthogonal gusset)
            AR=L23/(acHa+acHc)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARac2(
                    [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
               [ExcelArgument("Length of center rib")] double Lf,
                    [ExcelArgument("Gusset height at column face")] double hc,
               [ExcelArgument("Gusset height at free edge (<hc)")] double he,
               [ExcelArgument("Stiffener height (<he)")] double hs,
                    [ExcelArgument("Gusset width at beam face")] double wb,
               [ExcelArgument("Gusset width at free edge (<wb)")] double we,
               [ExcelArgument("Stiffener width (<we)")] double ws)
        {
            double L = L23(angle, Lf, hc, he, hs, wb, we, ws);
            double H = acHa(angle, Lf, hc, he, hs, wb, we, ws) + acHc(angle, Lf, hc, he, hs, wb, we, ws);

            return L / H;
        }
        [ExcelFunction(Description = @"Panel BD aspect ratio (assuming orthogonal gusset)
            AR=L34/(bdHb+bdHd)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARbd2(
                    [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
               [ExcelArgument("Length of center rib")] double Lf,
                    [ExcelArgument("Gusset height at column face")] double hc,
               [ExcelArgument("Gusset height at free edge (<hc)")] double he,
               [ExcelArgument("Stiffener height (<he)")] double hs,
                    [ExcelArgument("Gusset width at beam face")] double wb,
               [ExcelArgument("Gusset width at free edge (<wb)")] double we,
               [ExcelArgument("Stiffener width (<we)")] double ws)
        {
            double L = L34(angle, Lf, hc, he, hs, wb, we, ws);
            double H = bdHb(angle, Lf, hc, he, hs, wb, we, ws) + bdHd(angle, Lf, hc, he, hs, wb, we, ws);

            return L / H;
        }

        [ExcelFunction(Description = @"Panel A internal angle (assuming orthogonal gusset)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaA2(
                    [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
               [ExcelArgument("Length of center rib")] double Lf,
                    [ExcelArgument("Gusset height at column face")] double hc,
               [ExcelArgument("Gusset height at free edge (<hc)")] double he,
               [ExcelArgument("Stiffener height (<he)")] double hs,
                    [ExcelArgument("Gusset width at beam face")] double wb,
               [ExcelArgument("Gusset width at free edge (<wb)")] double we,
               [ExcelArgument("Stiffener width (<we)")] double ws)
        {
            double L = L23(angle, Lf, hc, he, hs, wb, we, ws);
            double Ha = abHa(angle, Lf, hc, he, hs, wb, we, ws);

            return Atan(Ha / Sqrt(Pow(L, 2) - Pow(Ha, 2)));
        }
        [ExcelFunction(Description = @"Panel B internal angle (assuming orthogonal gusset)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaB2(
                    [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
               [ExcelArgument("Length of center rib")] double Lf,
                    [ExcelArgument("Gusset height at column face")] double hc,
               [ExcelArgument("Gusset height at free edge (<hc)")] double he,
               [ExcelArgument("Stiffener height (<he)")] double hs,
                    [ExcelArgument("Gusset width at beam face")] double wb,
               [ExcelArgument("Gusset width at free edge (<wb)")] double we,
               [ExcelArgument("Stiffener width (<we)")] double ws)
        {
            double L = L34(angle, Lf, hc, he, hs, wb, we, ws);
            double Hb = abHb(angle, Lf, hc, he, hs, wb, we, ws);

            return Atan(Hb / Sqrt(Pow(L, 2) - Pow(Hb, 2)));
        }

        // Gusset panelization from panel dimensions
        [ExcelFunction(Description = @"Panel AB aspect ratio (assuming orthogonal gusset)
            AR=L13/(abHa+abHb)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARab3(
            [ExcelArgument("Length of yield line between panels A and B")] double L13,
            [ExcelArgument("Panel A dimension perpendicular to yield line between panels A and B")] double abHa,
            [ExcelArgument("Panel B dimension perpendicular to yield line between panels A and B")] double abHb)
        {
            return L13 / (abHa + abHb);
        }
        [ExcelFunction(Description = @"Panel AC aspect ratio (assuming orthogonal gusset)
            AR=L23/(acHa+acHc)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARac3(
            [ExcelArgument("Length of yield line between panels A and C")] double L23,
            [ExcelArgument("Panel A dimension perpendicular to yield line between panels A and C")] double acHa,
            [ExcelArgument("Panel C dimension perpendicular to yield line between panels A and C")] double acHc)
        {
            return L23 / (acHa + acHc);
        }
        [ExcelFunction(Description = @"Panel BD aspect ratio (assuming orthogonal gusset)
            AR=L34/(bdHb+bdHd)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_ARbd3(
            [ExcelArgument("Length of yield line between panels B and D")] double L34,
            [ExcelArgument("Panel B dimension perpendicular to yield line between panels B and D")] double bdHb,
            [ExcelArgument("Panel D dimension perpendicular to yield line between panels B and D")] double bdHd)
        {
            return L34 / (bdHb + bdHd);
        }

        [ExcelFunction(Description = @"Panel A internal angle (assuming orthogonal gusset)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaA3(
            [ExcelArgument("Length of yield line between panels A and C")] double L23,
            [ExcelArgument("Panel A dimension perpendicular to yield line between panels A and B")] double abHa)
        {
            return Atan(abHa / Sqrt(Pow(L23, 2) - Pow(abHa, 2)));
        }
        [ExcelFunction(Description = @"Panel B internal angle (assuming orthogonal gusset)
            Kinoshita et al 2008: Out of plane stiffness and yield strength of cruciform connection for BRB. AIJ 73(632)")]
        public static double Kinoshita_thetaB3(
            [ExcelArgument("Length of yield line between panels B and D")] double L34,
            [ExcelArgument("Panel B dimension perpendicular to yield line between panels A and B")] double abHb)
        {
            return Atan(abHb / Sqrt(Pow(L34, 2) - Pow(abHb, 2)));
        }

        // Helper functions for panel dimensions
        // refer Kinoshita (2008) Eqns A-1 to A-9//
        private static double L23(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double Ha = abHa(angle, Lf, hc, he, hs, wb, we, ws);
            return Sqrt(Pow(Lf - (he - hs) * Sin(angle), 2) + Pow(Ha, 2));
        }
        private static double L34(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double Hb = abHb(angle, Lf, hc, he, hs, wb, we, ws);
            return Sqrt(Pow(Lf - (we - ws) * Cos(angle), 2) + Pow(Hb, 2));
        }
        private static double abHa(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            return (hc - he) / 2 / Cos(angle) + (he - hs) * Cos(angle);
        }
        private static double abHb(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            // Kinoshita Eqn A-5 and A-5// incorrect.Refer AIJ C3.5.5g
            return (hc - he) / 2 / Cos(angle) + (we - ws) * Sin(angle);
        }
        private static double acHa(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double L = L23(angle, Lf, hc, he, hs, wb, we, ws);
            double Ha = abHa(angle, Lf, hc, he, hs, wb, we, ws);


            return Lf / L * Ha;
        }
        private static double acHc(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double L = L23(angle, Lf, hc, he, hs, wb, we, ws);


            if (hasStiffener(hs))
                return hs / L * ((wb - we) / 2 + Lf * Cos(angle));
            else
                // symmetrical to Kinoshita Eqn. A-9//
                return wb / L * ((hc + he) / 2 - Lf * Sin(angle));

        }
        private static double bdHb(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double L = L34(angle, Lf, hc, he, hs, wb, we, ws);
            double Hb = abHb(angle, Lf, hc, he, hs, wb, we, ws);


            return Lf / L * Hb;
        }
        private static double bdHd(double angle, double Lf, double hc, double he, double hs, double wb, double we, double ws)
        {
            double L = L34(angle, Lf, hc, he, hs, wb, we, ws);


            if (hasStiffener(ws))
                return ws / L * ((hc - he) / 2 + Lf * Sin(angle));
            else
                return hc / L * ((wb + we) / 2 - Lf * Cos(angle));

        }
        private static bool hasStiffener(double stiffenerDimension)
        {
            if (stiffenerDimension == 0)
                return false;
            else
                return true;
        }
        #endregion
        
        #region Lanning's Method (Gusset Type A)
        [ExcelFunction(Description = @"Gusset rotational stiffness
                Returns [∑Kθ,Kθ1,Kθ2.1,Kθ2.2]
            Lanning 2014: Using BRB on long span bridges near seismic faults. PhD UCSD")]
        public static double[] Lanning_Kg(
             [ExcelArgument("Gusset thickness")] double tg,
             [ExcelArgument("Splice width (ie Wl)")] double W,
             [ExcelArgument("Gusset average clear length to flange")] double Lavrg,
             [ExcelArgument("Projected splice root to beam flange")] double xmin1,
             [ExcelArgument("Projected splice tip to beam flange")] double xmax1,
             [ExcelArgument("Projected splice root to column flange")] double xmin2,
             [ExcelArgument("Projected splice tip to column flange")] double xmax2,
             [ExcelArgument("BRB angle from horizontal (ie β)", Name = "θ")] double θ,
             [ExcelArgument("Gusset taper angle at beam side (ie π/2-η), set to false if no plate [default = 0 rad]", Name = "φ1")] object _φ1,
             [ExcelArgument("Gusset taper angle at column side (ie π/2-η), set to false if no plate [default = 0 rad]", Name = "φ2")] object _φ2,
             [ExcelArgument("Youngs modulus [default =205000MPa]",Name ="E")]  object _E,
             [ExcelArgument("Shear modulus [default =205000MPa/2(1+0.3)",Name ="G")] object _G)
        {
            bool hasPlate1 = (_φ1 is bool && (bool)_φ1 == false) ? false : true;
            bool hasPlate2 = (_φ2 is bool && (bool)_φ2 == false) ? false : true;

            double φ1 = Check<double>(_φ1, 0);
            double φ2 = Check<double>(_φ2, 0);
            double E = Check<double>(_E, 205000);
            double G = Check<double>(_G, 205000/2/(1+0.3));

            double I = Pow(tg, 3) / 12;
            double J = Pow(tg, 3) / 3; // not defined by Lanning, perhaps eqn 6.12?

            double Kθ1 = Lanning_Kθ1(E * I,W,Lavrg); // middle section
            double Kθ2 = hasPlate1 ? Lanning_Kθ2(E * I, G * J, W, θ, φ1, xmin1,xmax1) : 0;
            double Kθ3 = hasPlate2 ? Lanning_Kθ2(E * I, G * J, W, PI / 2 - θ, φ2, xmin2, xmax2) : 0;

            return new double[] { Kθ1 + Kθ2 + Kθ3, Kθ1, Kθ2, Kθ3 };
        }
        internal static double Lanning_Kθ1(double EI, double W, double L)
        {
            return 4 * EI * W / L;
        }
        internal static double Lanning_Kθ2(double EI, double GJ, double W, double θ, double φ,double xmin,double xmax) { 
            double A = 12 * Pow(Cos(φ) / Sin(θ), 3) * Cos(θ);
            double B = 6 * Pow(Cos(φ) / Sin(θ), 2) * Cos(θ) * Sin(θ);
            double C = GJ / (EI) * Cos(φ) / Sin(θ) * Pow(Cos(θ), 2) * (1 - (φ + θ) / (0.5 * PI));
            double D = 6 * Pow(Cos(φ) / Sin(θ), 2) * ((φ + θ) / (0.5 * PI)) * Cos(θ);
            double F = 4 * Cos(φ) * Cos(θ) * ((φ + θ) / (0.5 * PI));

            return EI * (Log(xmax / xmin) * (A + B + C - D - F)
                + (xmin / xmax-1) * (2 * A + B - D)
                - (Pow(xmin / xmax, 2) - 1) * A / 2);
        }

        #endregion

        #region Lanning's Method - dimensions
        [ExcelFunction(Description = @"Zone 2 minimum x dimension
            Lanning 2014: Using BRB on long span bridges near seismic faults. PhD UCSD")]
        public static double Lanning_xmin(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Splice width")] double Ws,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Clear length of center rib")] double Lg,
                [ExcelArgument("True/False is beam (Zone 2.1) or column (Zone 2.2) side [default = TRUE]")] object _isBeamSide,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx)
        {
            double Lgx = Check<double>(_Lgx, 0);
            bool isBeamSide = Check<bool>(_isBeamSide, true);

            if(isBeamSide) return new LanningGusset(angle, Ws, Ls, Lg, Lgx).xmin1();
            else return new LanningGusset(angle, Ws, Ls, Lg, Lgx).xmin2();
        }
        [ExcelFunction(Description = @"Zone 2 maximum x dimension
            Lanning 2014: Using BRB on long span bridges near seismic faults. PhD UCSD")]
        public static double Lanning_xmax(
                [ExcelArgument("BRB angle to the horizontal (radians)")] double angle,
                [ExcelArgument("Splice width")] double Ws,
                [ExcelArgument("Splice length")] double Ls,
                [ExcelArgument("Clear length of center rib")] double Lg,
                [ExcelArgument("True/False is beam (Zone 2.1) or column (Zone 2.2) side [default = TRUE]")] object _isBeamSide,
                [ExcelArgument("Horizontal offset at flange to CL [default = 0]")] object _Lgx)
        {
            double Lgx = Check<double>(_Lgx, 0);
            bool isBeamSide = Check<bool>(_isBeamSide, true);

            if (isBeamSide) return new LanningGusset(angle, Ws, Ls, Lg, Lgx).xmax1();
            else return new LanningGusset(angle, Ws, Ls, Lg, Lgx).xmax2();
        }
        #endregion
    }
}