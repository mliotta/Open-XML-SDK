// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TREND function.
/// TREND(known_y's, known_x's, new_x's, const) - returns values along a linear trend.
/// If new_x's is omitted, it is assumed to be the same as known_x's.
/// If const is TRUE or omitted, the intercept is calculated normally; if FALSE, intercept is forced to 0.
/// For Phase 0: Simplified to handle single values only.
/// </summary>
public sealed class TrendFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrendFunction Instance = new();

    private TrendFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TREND";

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

        // Need at least 1 data point
        if (yValues.Count < 1)
        {
            return CellValue.Error("#N/A");
        }

        // Handle new_x's (args[2])
        var newX = xValues[0]; // Default to first x value
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Number)
            {
                newX = args[2].NumericValue;
            }
            else
            {
                // If new_x's is omitted or not a number, use known_x's
                newX = xValues[0];
            }
        }

        // Handle const parameter (args[3])
        var useIntercept = true;
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Boolean)
            {
                useIntercept = args[3].BoolValue;
            }
            else if (args[3].Type == CellValueType.Number)
            {
                useIntercept = args[3].NumericValue != 0;
            }
        }

        // If only one data point
        if (yValues.Count == 1)
        {
            if (useIntercept)
            {
                return CellValue.FromNumber(yValues[0]);
            }
            else
            {
                // Force through origin: slope = y/x
                if (xValues[0] == 0.0)
                {
                    return CellValue.Error("#DIV/0!");
                }

                var slope = yValues[0] / xValues[0];
                return CellValue.FromNumber(slope * newX);
            }
        }

        // Calculate slope and intercept
        double trendSlope;
        double trendIntercept;

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

            trendSlope = sumProduct / sumSquaresX;
            trendIntercept = meanY - (trendSlope * meanX);
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

            trendSlope = sumXY / sumXX;
            trendIntercept = 0.0;
        }

        // Calculate trend value
        var result = trendIntercept + (trendSlope * newX);

        return CellValue.FromNumber(result);
    }
}
