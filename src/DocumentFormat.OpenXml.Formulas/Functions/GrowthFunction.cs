// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GROWTH function.
/// GROWTH(known_y's, known_x's, new_x's, const) - calculates exponential growth.
/// Fits the data to y = b * m^x and returns predicted values.
/// If new_x's is omitted, it is assumed to be the same as known_x's.
/// If const is TRUE or omitted, b is calculated normally; if FALSE, b is forced to 1.
/// For Phase 0: Simplified to handle single values only.
/// </summary>
public sealed class GrowthFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GrowthFunction Instance = new();

    private GrowthFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GROWTH";

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
        }

        // Handle const parameter (args[3])
        var useConstant = true;
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Boolean)
            {
                useConstant = args[3].BoolValue;
            }
            else if (args[3].Type == CellValueType.Number)
            {
                useConstant = args[3].NumericValue != 0;
            }
        }

        // If only one data point
        if (yValues.Count == 1)
        {
            if (useConstant)
            {
                return CellValue.FromNumber(yValues[0]);
            }
            else
            {
                // Force b=1: y = m^x, so m = y^(1/x)
                if (xValues[0] == 0.0)
                {
                    return CellValue.Error("#DIV/0!");
                }

                var m = System.Math.Pow(yValues[0], 1.0 / xValues[0]);
                var result = System.Math.Pow(m, newX);
                return CellValue.FromNumber(result);
            }
        }

        // Exponential regression: y = b * m^x
        // Taking log: ln(y) = ln(b) + x*ln(m)
        // This is linear regression on ln(y) vs x
        var lnYValues = yValues.Select(y => System.Math.Log(y)).ToList();

        double slope; // This will be ln(m)
        double intercept; // This will be ln(b)

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
            intercept = meanLnY - (slope * meanX);
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
            intercept = 0.0;
        }

        // Calculate result: y = exp(intercept) * exp(slope * newX)
        var lnResult = intercept + (slope * newX);
        var growth = System.Math.Exp(lnResult);

        return CellValue.FromNumber(growth);
    }
}
