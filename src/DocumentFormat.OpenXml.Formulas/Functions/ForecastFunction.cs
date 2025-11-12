// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FORECAST function.
/// FORECAST(x, known_y's, known_x's) - calculates a future value using linear regression.
/// Formula: y = a + bx, where b is the slope and a is the intercept.
/// </summary>
public sealed class ForecastFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ForecastFunction Instance = new();

    private ForecastFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FORECAST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Propagate errors
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

        // Get x value for prediction
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var x = args[0].NumericValue;

        var yValues = new List<double>();
        var xValues = new List<double>();

        // Extract numeric values from known_y's
        if (args[1].Type == CellValueType.Number)
        {
            yValues.Add(args[1].NumericValue);
        }

        // Extract numeric values from known_x's
        if (args[2].Type == CellValueType.Number)
        {
            xValues.Add(args[2].NumericValue);
        }

        // Arrays must have same length
        if (yValues.Count != xValues.Count)
        {
            return CellValue.Error("#N/A");
        }

        // Need at least 1 data point (Excel allows single point forecasting)
        if (yValues.Count < 1)
        {
            return CellValue.Error("#N/A");
        }

        // If only one data point, return its y value
        if (yValues.Count == 1)
        {
            return CellValue.FromNumber(yValues[0]);
        }

        // Calculate means
        var meanX = xValues.Average();
        var meanY = yValues.Average();

        // Calculate slope and intercept
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
        var intercept = meanY - (slope * meanX);

        // Calculate forecast: y = a + bx
        var forecast = intercept + (slope * x);

        return CellValue.FromNumber(forecast);
    }
}
