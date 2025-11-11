// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MODE function.
/// MODE(number1, [number2], ...) - returns most frequent value.
/// </summary>
public sealed class ModeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ModeFunction Instance = new();

    private ModeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MODE";

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
            return CellValue.Error("#N/A");
        }

        // Group by value and find the most frequent
        var groups = values.GroupBy(v => v)
            .OrderByDescending(g => g.Count())
            .ThenBy(g => g.Key);

        var mostFrequent = groups.FirstOrDefault();

        if (mostFrequent == null || mostFrequent.Count() < 2)
        {
            // MODE requires at least one value to appear more than once
            return CellValue.Error("#N/A");
        }

        return CellValue.FromNumber(mostFrequent.Key);
    }
}
