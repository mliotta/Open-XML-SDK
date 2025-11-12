// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOGEST function.
/// LOGEST(known_y's, known_x's, const, stats) - returns exponential regression statistics.
/// Fits the exponential curve y = b * m^x to the data.
/// In full Excel, this returns an array of statistics. For Phase 0, we return m (the base) only.
/// const: TRUE (default) = calculate b normally; FALSE = force b to 1
/// stats: TRUE = return additional statistics; FALSE (default) = return only m and b
/// Phase 0 returns only the m value.
/// </summary>
public sealed class LogestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LogestFunction Instance = new();

    private LogestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOGEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Propagate errors in known_y's
        if (args[0].IsError)
        {
            return args[0];
        }

        var yValues = new List<double>();
        var xValues = new List<double>();

        // Extract numeric values from known_y's
        if (args[0].Type == CellValueType.Number)
        {
            // All y values must be positive for exponential regression
            if (args[0].NumericValue <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            yValues.Add(args[0].NumericValue);
        }

        // Handle known_x's (args[1])
        if (args.Length >= 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type == CellValueType.Number)
            {
                xValues.Add(args[1].NumericValue);
            }
        }
        else
        {
            // If known_x's is omitted, use 1, 2, 3, ...
            for (int i = 0; i < yValues.Count; i++)
            {
                xValues.Add(i + 1);
            }
        }

        // Arrays must have same length
        if (yValues.Count != xValues.Count)
        {
            return CellValue.Error("#N/A");
        }

        // Need at least 2 data points for regression
        if (yValues.Count < 2)
        {
            return CellValue.Error("#N/A");
        }

        // Handle const parameter (args[2])
        var useConstant = true;
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Boolean)
            {
                useConstant = args[2].BoolValue;
            }
            else if (args[2].Type == CellValueType.Number)
            {
                useConstant = args[2].NumericValue != 0;
            }
        }

        // Handle stats parameter (args[3]) - currently ignored in Phase 0

        // Exponential regression: y = b * m^x
        // Taking log: ln(y) = ln(b) + x*ln(m)
        // This is linear regression on ln(y) vs x
        var lnYValues = yValues.Select(y => System.Math.Log(y)).ToList();

        double slope; // This will be ln(m)

        if (useConstant)
        {
            // Normal exponential regression
            var meanX = xValues.Average();
            var meanLnY = lnYValues.Average();

            var sumProduct = 0.0;
            var sumSquaresX = 0.0;

            for (int i = 0; i < xValues.Count; i++)
            {
                var diffX = xValues[i] - meanX;
                var diffLnY = lnYValues[i] - meanLnY;

                sumProduct += diffX * diffLnY;
                sumSquaresX += diffX * diffX;
            }

            if (sumSquaresX == 0.0)
            {
                return CellValue.Error("#DIV/0!");
            }

            slope = sumProduct / sumSquaresX;
        }
        else
        {
            // Force b=1 (ln(b)=0): ln(y) = x*ln(m)
            // slope = Σ(x*ln(y)) / Σ(x²)
            var sumXLnY = 0.0;
            var sumXX = 0.0;

            for (int i = 0; i < xValues.Count; i++)
            {
                sumXLnY += xValues[i] * lnYValues[i];
                sumXX += xValues[i] * xValues[i];
            }

            if (sumXX == 0.0)
            {
                return CellValue.Error("#DIV/0!");
            }

            slope = sumXLnY / sumXX;
        }

        // Convert slope (ln(m)) back to m
        var m = System.Math.Exp(slope);

        // Phase 0: Return only m (the exponential base)
        // Full implementation would return an array with m, b, and optionally more stats
        return CellValue.FromNumber(m);
    }
}
