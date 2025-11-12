// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AVEDEV function.
/// AVEDEV(number1, [number2], ...) - returns the average of the absolute deviations of data points from their mean.
/// </summary>
public sealed class AvedevFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AvedevFunction Instance = new();

    private AvedevFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AVEDEV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

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

        if (values.Count == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate mean
        double sum = 0.0;
        foreach (var value in values)
        {
            sum += value;
        }
        double mean = sum / values.Count;

        // Calculate average of absolute deviations
        double deviationSum = 0.0;
        foreach (var value in values)
        {
            deviationSum += System.Math.Abs(value - mean);
        }

        double result = deviationSum / values.Count;
        return CellValue.FromNumber(result);
    }
}
