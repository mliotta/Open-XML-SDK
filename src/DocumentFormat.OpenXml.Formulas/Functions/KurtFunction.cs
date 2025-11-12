// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the KURT function.
/// KURT(number1, [number2], ...) - Returns the kurtosis of a distribution (tailedness measure).
/// </summary>
public sealed class KurtFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly KurtFunction Instance = new();

    private KurtFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "KURT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var values = new List<double>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                values.Add(arg.NumericValue);
            }
        }

        // KURT requires at least 4 data points
        if (values.Count < 4)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate mean
        var mean = values.Average();

        // Calculate standard deviation (sample)
        var n = values.Count;
        var sumSquaredDiffs = values.Sum(v => System.Math.Pow(v - mean, 2));
        var variance = sumSquaredDiffs / (n - 1);
        var stdev = System.Math.Sqrt(variance);

        // If standard deviation is zero, kurtosis is undefined
        if (stdev == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate excess kurtosis using Excel's formula
        // Kurt = (n(n+1)/((n-1)(n-2)(n-3))) * Σ((x-x̄)/s)⁴ - 3(n-1)²/((n-2)(n-3))
        var sumFourthPowerZScores = values.Sum(v => System.Math.Pow((v - mean) / stdev, 4));

        var term1 = (n * (n + 1.0)) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * sumFourthPowerZScores;
        var term2 = (3.0 * System.Math.Pow(n - 1.0, 2)) / ((n - 2.0) * (n - 3.0));
        var kurtosis = term1 - term2;

        return CellValue.FromNumber(kurtosis);
    }
}
