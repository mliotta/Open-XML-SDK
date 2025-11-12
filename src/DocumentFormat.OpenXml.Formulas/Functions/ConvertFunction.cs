// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONVERT function.
/// CONVERT(number, from_unit, to_unit) - converts number between units.
/// </summary>
public sealed class ConvertFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConvertFunction Instance = new();

    private ConvertFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONVERT";

    private sealed class UnitInfo
    {
        public string Category { get; set; }

        public double Factor { get; set; }

        public double Offset { get; set; }

        public UnitInfo(string category, double factor, double offset)
        {
            Category = category;
            Factor = factor;
            Offset = offset;
        }
    }

    // Conversion factors to base units
    private static readonly Dictionary<string, UnitInfo> _units = new Dictionary<string, UnitInfo>(StringComparer.OrdinalIgnoreCase)
    {
        // Weight - base unit: gram
        { "g", new UnitInfo("Weight", 1.0, 0.0) },
        { "kg", new UnitInfo("Weight", 1000.0, 0.0) },
        { "lbm", new UnitInfo("Weight", 453.59237, 0.0) },
        { "ozm", new UnitInfo("Weight", 28.349523125, 0.0) },

        // Distance - base unit: meter
        { "m", new UnitInfo("Distance", 1.0, 0.0) },
        { "km", new UnitInfo("Distance", 1000.0, 0.0) },
        { "mi", new UnitInfo("Distance", 1609.344, 0.0) },
        { "ft", new UnitInfo("Distance", 0.3048, 0.0) },
        { "in", new UnitInfo("Distance", 0.0254, 0.0) },

        // Time - base unit: second
        { "sec", new UnitInfo("Time", 1.0, 0.0) },
        { "min", new UnitInfo("Time", 60.0, 0.0) },
        { "hr", new UnitInfo("Time", 3600.0, 0.0) },
        { "day", new UnitInfo("Time", 86400.0, 0.0) },
        { "yr", new UnitInfo("Time", 31557600.0, 0.0) }, // 365.25 days

        // Temperature - special handling required
        { "C", new UnitInfo("Temperature", 1.0, 0.0) },      // Celsius (base)
        { "F", new UnitInfo("Temperature", 0.0, 0.0) },      // Fahrenheit
        { "K", new UnitInfo("Temperature", 1.0, 273.15) },   // Kelvin

        // Volume - base unit: liter
        { "l", new UnitInfo("Volume", 1.0, 0.0) },
        { "gal", new UnitInfo("Volume", 3.785411784, 0.0) },
        { "qt", new UnitInfo("Volume", 0.946352946, 0.0) },
        { "pt", new UnitInfo("Volume", 0.473176473, 0.0) },
    };

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[2].IsError)
        {
            return args[2];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var fromUnit = args[1].StringValue;
        var toUnit = args[2].StringValue;

        if (!_units.TryGetValue(fromUnit, out var fromInfo))
        {
            return CellValue.Error("#N/A");
        }

        if (!_units.TryGetValue(toUnit, out var toInfo))
        {
            return CellValue.Error("#N/A");
        }

        if (fromInfo.Category != toInfo.Category)
        {
            return CellValue.Error("#N/A");
        }

        double result;

        // Special handling for temperature conversions
        if (fromInfo.Category == "Temperature")
        {
            result = ConvertTemperature(number, fromUnit, toUnit);
        }
        else
        {
            // Convert to base unit, then to target unit
            var baseValue = number * fromInfo.Factor + fromInfo.Offset;
            result = (baseValue - toInfo.Offset) / toInfo.Factor;
        }

        return CellValue.FromNumber(result);
    }

    private static double ConvertTemperature(double value, string fromUnit, string toUnit)
    {
        // Convert to Celsius first
        double celsius = fromUnit.ToUpperInvariant() switch
        {
            "C" => value,
            "F" => (value - 32.0) * 5.0 / 9.0,
            "K" => value - 273.15,
            _ => value
        };

        // Convert from Celsius to target unit
        return toUnit.ToUpperInvariant() switch
        {
            "C" => celsius,
            "F" => celsius * 9.0 / 5.0 + 32.0,
            "K" => celsius + 273.15,
            _ => celsius
        };
    }
}
