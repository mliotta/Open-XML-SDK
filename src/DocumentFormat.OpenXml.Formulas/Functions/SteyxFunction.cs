// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the STEYX function.
/// STEYX(known_y's, known_x's) - returns the standard error of the predicted y-value for each x in regression.
/// </summary>
public sealed class SteyxFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SteyxFunction Instance = new();

    private SteyxFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "STEYX";

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
        if (n < 3)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate means
        double xSum = 0, ySum = 0;
        for (int i = 0; i < n; i++)
        {
            xSum += xValues[i];
            ySum += yValues[i];
        }
        double xMean = xSum / n;
        double yMean = ySum / n;

        // Calculate slope and intercept
        double numerator = 0, denominator = 0;
        for (int i = 0; i < n; i++)
        {
            double xDev = xValues[i] - xMean;
            double yDev = yValues[i] - yMean;
            numerator += xDev * yDev;
            denominator += xDev * xDev;
        }

        if (denominator == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        double slope = numerator / denominator;
        double intercept = yMean - slope * xMean;

        // Calculate standard error
        double sumSquaredErrors = 0;
        for (int i = 0; i < n; i++)
        {
            double predicted = intercept + slope * xValues[i];
            double error = yValues[i] - predicted;
            sumSquaredErrors += error * error;
        }

        double standardError = System.Math.Sqrt(sumSquaredErrors / (n - 2));

        return CellValue.FromNumber(standardError);
    }
}
