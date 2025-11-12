// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RSQ function.
/// RSQ(known_y's, known_x's) - returns the square of the Pearson correlation coefficient (R-squared).
/// </summary>
public sealed class RsqFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RsqFunction Instance = new();

    private RsqFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RSQ";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // First half are y values, second half are x values
        int midpoint = args.Length / 2;
        var yValues = new List<double>();
        var xValues = new List<double>();

        // Collect y values
        for (int i = 0; i < midpoint; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                yValues.Add(args[i].NumericValue);
            }
        }

        // Collect x values
        for (int i = midpoint; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                xValues.Add(args[i].NumericValue);
            }
        }

        if (xValues.Count == 0 || yValues.Count == 0 || xValues.Count != yValues.Count)
        {
            return CellValue.Error("#N/A");
        }

        int n = xValues.Count;

        // Calculate means
        double xSum = 0, ySum = 0;
        for (int i = 0; i < n; i++)
        {
            xSum += xValues[i];
            ySum += yValues[i];
        }
        double xMean = xSum / n;
        double yMean = ySum / n;

        // Calculate R-squared using correlation coefficient
        double numerator = 0, xDenom = 0, yDenom = 0;
        for (int i = 0; i < n; i++)
        {
            double xDev = xValues[i] - xMean;
            double yDev = yValues[i] - yMean;
            numerator += xDev * yDev;
            xDenom += xDev * xDev;
            yDenom += yDev * yDev;
        }

        if (xDenom == 0 || yDenom == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        double r = numerator / System.Math.Sqrt(xDenom * yDenom);
        double rSquared = r * r;

        return CellValue.FromNumber(rSquared);
    }
}
