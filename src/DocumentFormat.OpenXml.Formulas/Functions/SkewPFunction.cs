// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SKEW.P function.
/// SKEW.P(number1, [number2], ...) - returns the population skewness of a distribution.
/// </summary>
public sealed class SkewPFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SkewPFunction Instance = new();

    private SkewPFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SKEW.P";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var values = new List<double>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }

            if (arg.Type == CellValueType.Number)
            {
                values.Add(arg.NumericValue);
            }
        }

        // SKEW.P requires at least 3 data points
        if (values.Count < 3)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate mean
        var mean = values.Average();

        // Calculate standard deviation (population)
        var n = values.Count;
        var sumSquaredDiffs = values.Sum(v => System.Math.Pow(v - mean, 2));
        var variance = sumSquaredDiffs / n;
        var stdev = System.Math.Sqrt(variance);

        // If standard deviation is zero, skewness is undefined
        if (stdev == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate population skewness
        // Skew.P = (1/n) * Σ((x-x̄)/s)³
        var sumCubedZScores = values.Sum(v => System.Math.Pow((v - mean) / stdev, 3));
        var skewness = sumCubedZScores / n;

        return CellValue.FromNumber(skewness);
    }
}
