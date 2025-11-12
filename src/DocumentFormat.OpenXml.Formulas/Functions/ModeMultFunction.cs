// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MODE.MULT function.
/// MODE.MULT(number1, [number2], ...) - returns a vertical array of the most frequently occurring values.
/// Note: This returns the smallest mode value when multiple modes exist (simplified implementation).
/// </summary>
public sealed class ModeMultFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ModeMultFunction Instance = new();

    private ModeMultFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MODE.MULT";

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
            .ThenBy(g => g.Key)
            .ToList();

        // Find the maximum frequency
        var maxFrequency = groups.FirstOrDefault()?.Count() ?? 0;

        if (maxFrequency < 2)
        {
            // MODE requires at least one value to appear more than once
            return CellValue.Error("#N/A");
        }

        // Get all values with the maximum frequency
        var modes = groups
            .Where(g => g.Count() == maxFrequency)
            .Select(g => g.Key)
            .OrderBy(v => v)
            .ToList();

        // For simplicity, return the first (smallest) mode
        // In Excel, this would return a vertical array
        // Full array support would require returning multiple values
        return CellValue.FromNumber(modes.First());
    }
}
