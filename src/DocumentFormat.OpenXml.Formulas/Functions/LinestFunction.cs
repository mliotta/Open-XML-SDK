// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LINEST function.
/// LINEST(known_y's, known_x's, const, stats) - returns linear regression statistics.
/// In full Excel, this returns an array of statistics. For Phase 0, we return the slope only.
/// const: TRUE (default) = calculate intercept normally; FALSE = force intercept to 0
/// stats: TRUE = return additional statistics; FALSE (default) = return only slope and intercept
/// Phase 0 returns only the slope value.
/// </summary>
public sealed class LinestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LinestFunction Instance = new();

    private LinestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LINEST";

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
        var useIntercept = true;
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Boolean)
            {
                useIntercept = args[2].BoolValue;
            }
            else if (args[2].Type == CellValueType.Number)
            {
                useIntercept = args[2].NumericValue != 0;
            }
        }

        // Handle stats parameter (args[3]) - currently ignored in Phase 0
        // In full implementation, this would determine whether to return additional statistics

        // Calculate slope
        double slope;

        if (useIntercept)
        {
            // Normal linear regression
            var meanX = xValues.Average();
            var meanY = yValues.Average();

            var sumProduct = 0.0;
            var sumSquaresX = 0.0;

            for (int i = 0; i < xValues.Count; i++)
            {
                var diffX = xValues[i] - meanX;
                var diffY = yValues[i] - meanY;

                sumProduct += diffX * diffY;
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
            // Force intercept to 0: slope = Σ(xy) / Σ(x²)
            var sumXY = 0.0;
            var sumXX = 0.0;

            for (int i = 0; i < xValues.Count; i++)
            {
                sumXY += xValues[i] * yValues[i];
                sumXX += xValues[i] * xValues[i];
            }

            if (sumXX == 0.0)
            {
                return CellValue.Error("#DIV/0!");
            }

            slope = sumXY / sumXX;
        }

        // Phase 0: Return only the slope
        // Full implementation would return an array with slope, intercept, and optionally more stats
        return CellValue.FromNumber(slope);
    }
}
