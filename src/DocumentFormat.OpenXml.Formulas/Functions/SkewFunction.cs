// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SKEW function.
/// SKEW(number1, [number2], ...) - Returns the skewness of a distribution (asymmetry measure).
/// </summary>
public sealed class SkewFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SkewFunction Instance = new();

    private SkewFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SKEW";

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

        // SKEW requires at least 3 data points
        if (values.Count < 3)
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

        // If standard deviation is zero, skewness is undefined
        if (stdev == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate skewness using the sample skewness formula
        // Skew = (n/((n-1)(n-2))) * Σ((x-x̄)/s)³
        var sumCubedZScores = values.Sum(v => System.Math.Pow((v - mean) / stdev, 3));
        var skewness = (n / ((n - 1.0) * (n - 2.0))) * sumCubedZScores;

        return CellValue.FromNumber(skewness);
    }
}
