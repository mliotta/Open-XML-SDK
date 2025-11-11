// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the STDEVP function.
/// STDEVP(number1, [number2], ...) - standard deviation (population).
/// </summary>
public sealed class StDevPFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly StDevPFunction Instance = new();

    private StDevPFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "STDEVP";

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

        if (values.Count == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate mean
        var mean = values.Average();

        // Calculate variance (population)
        var sumSquaredDiffs = values.Sum(v => System.Math.Pow(v - mean, 2));
        var variance = sumSquaredDiffs / values.Count;

        // Standard deviation is square root of variance
        var stdev = System.Math.Sqrt(variance);

        return CellValue.FromNumber(stdev);
    }
}
