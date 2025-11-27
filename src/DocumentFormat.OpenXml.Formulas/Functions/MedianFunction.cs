// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MEDIAN function.
/// MEDIAN(number1, [number2], ...) - returns median value.
/// </summary>
public sealed class MedianFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MedianFunction Instance = new();

    private MedianFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MEDIAN";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        var values = new List<double>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == FormulaResultType.Number)
            {
                values.Add(arg.NumericValue);
            }
        }

        if (values.Count == 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        values.Sort();

        var count = values.Count;
        var middle = count / 2;

        if (count % 2 == 0)
        {
            // Even number of values - average the two middle values
            var median = (values[middle - 1] + values[middle]) / 2.0;
            return FormulaResult.FromNumber(median);
        }
        else
        {
            // Odd number of values - take the middle value
            return FormulaResult.FromNumber(values[middle]);
        }
    }
}
