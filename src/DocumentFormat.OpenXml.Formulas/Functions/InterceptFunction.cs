// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the INTERCEPT function.
/// INTERCEPT(known_y's, known_x's) - Returns the y-intercept of the linear regression line.
/// </summary>
public sealed class InterceptFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly InterceptFunction Instance = new();

    private InterceptFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "INTERCEPT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        var yValues = new List<double>();
        var xValues = new List<double>();

        // Extract numeric values from known_y's
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type == CellValueType.Number)
        {
            yValues.Add(args[0].NumericValue);
        }

        // Extract numeric values from known_x's
        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[1].Type == CellValueType.Number)
        {
            xValues.Add(args[1].NumericValue);
        }

        // Arrays must have same length
        if (yValues.Count != xValues.Count)
        {
            return CellValue.Error("#N/A");
        }

        // Need at least 2 data points
        if (yValues.Count < 2)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate means
        var meanX = xValues.Average();
        var meanY = yValues.Average();

        // Calculate slope first
        // Slope = Σ((x-x̄)(y-ȳ)) / Σ(x-x̄)²
        var sumProduct = 0.0;
        var sumSquaresX = 0.0;

        for (int i = 0; i < xValues.Count; i++)
        {
            var diffX = xValues[i] - meanX;
            var diffY = yValues[i] - meanY;

            sumProduct += diffX * diffY;
            sumSquaresX += diffX * diffX;
        }

        // Check for division by zero (no variance in x)
        if (sumSquaresX == 0.0)
        {
            return CellValue.Error("#DIV/0!");
        }

        var slope = sumProduct / sumSquaresX;

        // Calculate intercept
        // Intercept = ȳ - slope × x̄
        var intercept = meanY - (slope * meanX);

        return CellValue.FromNumber(intercept);
    }
}
