using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Math;
namespace BRBBuckling
{
    internal class ThorntonGusset
    {
        double θ; // BRB angle
        Point p1; // midpoint
        Point p2; // lower point
        Point p3; // upper point

        public ThorntonGusset(double θ, double W, double L,
             double Lg, double Lgx = 0)
        {
            // BRB angle
            this.θ = θ;

            // centerline control points
            var origin = new Point(0, 0);
            p1 = origin.Move(Lgx, 0).Move(Lg, θ);

            Line beamFlange = new Line(origin, 0);
            Line colFlange = new Line(origin, PI / 2);
            Line stiffenerBase = new Line(p1, θ - PI);

            double Wb = L * Tan(30 * PI / 180);

            p2 = p1.Move(Wb + W / 2, θ - PI / 2);
            p3 = p1.Move(Wb + W / 2, θ + PI / 2);

            if (p2.y <= 0) p2 = beamFlange.Intersect(stiffenerBase);
            if (p3.x <= 0) p3 = colFlange.Intersect(stiffenerBase);
        }

        // Whitmore width & Thornton length
        public double W() { return p2.DistanceTo(p3); }
        public double L1() { return Distance(p1); }
        public double L2() { return Distance(p2); }
        public double L3() { return Distance(p3); }

        private double Distance(Point p)
        {
            Point p1 = new Line(0).Intersect(new Line(p, this.θ));
            Point p2 = new Line(PI / 2).Intersect(new Line(p, this.θ));
            return Min(p1.DistanceTo(p), p2.DistanceTo(p));
        }
    }

    internal class KinoshitaGusset
    {
        double θ; /* BRB angle */
        public Point p0; // origin
        public Point p1; /* tip */ public Point p3; /* stiffener root */
        public Point p2; /* beam edge */ public Point p5; /* beam flange */
        public Point p4; /* column edge */ public Point p6; /* column flange */

        /// <param name="θ">BRB angle</param>
        /// <param name="W1">gusset tip half width (CL to horiz edge)</param>
        /// <param name="W2">gusset tip half width (CL to vert edge)</param>
        /// <param name="L">Splice length</param>
        /// <param name="Lg">Clear gusset length to splice along CL</param>
        /// <param name="Lgx">Horizonal offset to CL</param>
        /// <param name="horizStiffenerPerc">Horizontal edge stiffener percentage length</param>
        /// <param name="vertStiffenerPerc">Vertical edge stiffener percentage length</param>
        /// <param name="horizTaperAngle">horizontal edge taper (+ve reduces gusset size)</param>
        /// <param name="vertTaperAngle">vertical edge taper (+ve reduces gusset size)</param>
        public KinoshitaGusset(double θ, double W1,double W2, double L,
             double Lg, double Lgx = 0,
             double horizStiffenerPerc = 0, double vertStiffenerPerc = 0,
             double horizTaperAngle = 0, double vertTaperAngle = 0)
        {
            // BRB angle
            this.θ = θ;

            // centerline control points
            p0 = new Point(0, 0);
            p3 = p0.Move(Lgx, 0).Move(Lg, θ);
            p1 = p3.Move(L, θ);

            // column edge control points
            Point p1a = p1.Move(W1, θ + PI / 2);
            Line colFlange = new Line(p0, PI / 2);
            Line colEdge = new Line(p1a, horizTaperAngle + PI);
            if (horizStiffenerPerc == 0)
            {
                p6 = p0;
                p4 = colFlange.Intersect(colEdge);
            }
            else
            {
                p6 = colFlange.Intersect(colEdge);
                p4 = p6.Move(p6.DistanceTo(p1a) * horizStiffenerPerc,
                     PI / 2 - horizTaperAngle);
            }
            // beam edge control points
            Point p1b = p1.Move(W2, θ - PI / 2);
            Line beamFlange = new Line(p0, 0);
            Line beamEdge = new Line(p1b, -PI / 2 - vertTaperAngle);
            if (vertStiffenerPerc == 0)
            {
                p5 = p0;
                p2 = beamFlange.Intersect(beamEdge);
            }
            else
            {
                p5 = beamFlange.Intersect(beamEdge);
                p2 = p5.Move(p5.DistanceTo(p1b) * vertStiffenerPerc,
                     PI / 2 - vertTaperAngle);
            }

        }

        // Kinoshita panel dimensions
        public double ARab() { return AspectRatio(p3, p4, p1, p2); }
        public double ARac() { return AspectRatio(p3, p5, p2, p1); }
        public double ARbd() { return AspectRatio(p3, p6, p4, p1); }
        public double ThetaA() { return p3.Angle(p1, p2); }
        public double ThetaB() { return p3.Angle(p1, p4); }
        public double PhiA() { return PI/2-p1.Angle(p3, p2); }
        public double PhiB() { return PI/2-p1.Angle(p3, p4); }
        /// <summary>panel aspect ratio with yield line from p1 to p3</summary>
        public double AspectRatio(Point p1, Point p2, Point p3, Point p4)
        {
            var yieldLine = new Line(p1, p3);
            return p1.DistanceTo(p3) /
                 (yieldLine.DistanceTo(p2) + yieldLine.DistanceTo(p4));
        }
    }

    internal class LanningGusset
    {
        double θ; // BRB angle
        double L;
        Point p; // midpoint
        Point p1; // vertical edge splice corner
        Point p2; // horiz edge splice corner

        public LanningGusset(double θ, double W, double L,
             double Lg, double Lgx = 0)
        {
            // BRB angle
            this.θ = θ;
            this.L = L;

            // centerline control points
            var origin = new Point(0, 0);
            p = origin.Move(Lgx, 0).Move(Lg, θ);

            p1 = p.Move(W / 2, θ - PI / 2); // vertical edge
            p2 = p.Move(W / 2, θ + PI / 2); // horiz edge
        }

        // Whitmore width & Thornton length
        public double W() { return p1.DistanceTo(p2); }
        public double xmin1() { return PositiveDistance(p1); }
        public double xmax1() { return PositiveDistance(p1.Move(L, θ)); }
        public double xmin2() { return PositiveDistance(p2); }
        public double xmax2() { return PositiveDistance(p2.Move(L, θ)); }

        private double PositiveDistance(Point pt)
        {
            // beam flange (vert edge)
            Point p1a = new Line(0).Intersect(new Line(pt, this.θ));
            double d1 = pt.y < 0 ? 0 : p1a.DistanceTo(pt);

            // column flange (horiz edge)
            Point p2a = new Line(PI / 2).Intersect(new Line(pt, this.θ));
            double d2 = pt.x < 0 ? 0 : p2a.DistanceTo(pt);

            return Min(d1, d2);
        }
    }

    public struct Point
    {
        public double x; public double y;
        public Point(double x, double y) { this.x = x; this.y = y; }
        public double DistanceTo(Point pt) { return Sqrt(Pow(pt.x - this.x, 2) + Pow(pt.y - y, 2)); }
        public Point Move(double d, double angle)
        {
            if (angle == 0) return new Point(x + d, y);
            else if (angle == PI / 2) return new Point(x, y + d);
            else if (Abs(angle) == PI) return new Point(x - d, y);
            else if (angle == -PI / 2 || angle == 3 * PI / 2) return new Point(x, y - d);
            else return new Point(x + d * Cos(angle), y + d * Sin(angle));
        }
        /// <summary>included angle</summary>
        public double Angle(Point p1, Point p2)
        {
            var d1 = this.DistanceTo(p1);
            var d2 = this.DistanceTo(p2);
            var d12 = p1.DistanceTo(p2);
            return Acos((Pow(d1, 2) + Pow(d2, 2) - Pow(d12, 2)) / (2 * d1 * d2));
        }
        /// <summary>angle from horizontal</summary>
        public double Angle(Point p)
        {
            Point p2 = new Point(this.x+10, this.y);
            return this.Angle(p, p2);
        }

    }
    public struct Line
    {
        public double m; public double b;
        public Line(Point p1, Point p2)
        {
            this.m = (p2.y - p1.y) / (p2.x - p1.x);
            this.b = p1.y - m * p1.x;
        }
        public Line(double angle)
        {
            this.m = Tan(angle); this.b = 0.0;
        }
        public Line(Point p, double angle)
        {
            this.m = Tan(angle); this.b = p.y - m * p.x;
        }
        public Point Intersect(Line other)
        {
            double x = (other.b - b) / (m - other.m);
            double y = m * x + b;
            return new Point(x, y);
        }
        public double DistanceTo(Point p)
        {
            return Abs(m * p.x + b - p.y) / Sqrt(m * m + 1);
        }
    }
}