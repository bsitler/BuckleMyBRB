using System;
using ExcelDna.Integration;
using System.Collections.Generic;

public static class UnitConversionFunctions
{
    [ExcelFunction(Description = @"Unit conversion (case insensitive)
        Force:    N,kN,MN,kip
        Length:    mm,cm,m,in,ft,yd,mile
        Moment:  Nm,kNm,ft_lb,kip_in
        Area:        mm²,cm²,m²,in²,ft²
        Stress:    kPa,MPa,GPa,psi,psf,ksi")]
    public static double Unit(
        [ExcelArgument("Value")] double value, 
        [ExcelArgument("Previous unit",Name ="previous_Unit")] string _prevUnit, 
        [ExcelArgument("New unit",Name ="new_Unit")] string _newUnit)
    {
        double prevUnit; double newUnit;
        if (ForceUnits.TryGetValue(_prevUnit, out prevUnit)
            && ForceUnits.TryGetValue(_newUnit, out newUnit))
            return value * newUnit / prevUnit;
        else if (LengthUnits.TryGetValue(_prevUnit, out prevUnit)
            && LengthUnits.TryGetValue(_newUnit, out newUnit))
            return value * newUnit / prevUnit;
        else if (AreaUnits.TryGetValue(_prevUnit, out prevUnit)
            && AreaUnits.TryGetValue(_newUnit, out newUnit))
            return value * newUnit / prevUnit;
        else if (MomentUnits.TryGetValue(_prevUnit, out prevUnit)
            && MomentUnits.TryGetValue(_newUnit, out newUnit))
            return value * newUnit / prevUnit;
        else if (PressureUnits.TryGetValue(_prevUnit, out prevUnit)
            && PressureUnits.TryGetValue(_newUnit, out newUnit))
            return value * newUnit / prevUnit;
        else return double.NaN;
    }

    // Unit names are case insensitive, so make sure there are no duplicates
    static Dictionary<string, double> ForceUnits = 
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
        {"N", 1000},
        { "kN",1},
        { "MN",0.001},
        { "kip", 0.22480894}};
    static Dictionary<string, double> LengthUnits =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
        { "m", 1.0},
        { "cm",100},
        { "mm",1000},
        { "in",39.37008},
        { "ft",3.28084},
        { "yd",1.09361333333333},
        { "mile",0.000621371212121212}};
    static Dictionary<string, double> AreaUnits =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
        {"mm2", 1000000},
        { "cm2",10000},
        { "m2",1},
        { "in2",1550},
        { "ft2",10.7638888888889}};
    static Dictionary<string, double> MomentUnits =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
        {"Nm", 1000},
        { "kNm",1},
        { "ft_lb",737.5621212},
        { "kip_in",8.8507454544}};
    static Dictionary<string, double> PressureUnits =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase) {
        { "kPa", 1},
        { "MPa",0.001},
        { "GPa",0.000001},
        { "psi",0.145038},
        { "psf",20.885472},
        { "ksi",0.000145038}};
}