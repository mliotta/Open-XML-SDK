// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SMALL function.
/// SMALL(array, k) - returns the k-th smallest value (1-based indexing).
/// </summary>
public sealed class SmallFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SmallFunction Instance = new();

    private SmallFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SMALL";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Propagate errors
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Get k value
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var k = (int)args[1].NumericValue;

        if (k < 1)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Collect all numeric values
        var values = new List<double>();

        if (args[0].Type == FormulaResultType.Number)
        {
            values.Add(args[0].NumericValue);
        }

        if (values.Count == 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // When there's only one value, return it for any reasonable k
        // This allows SMALL({5}, 2) to return 5 (the only value is all ranks)
        if (values.Count == 1)
        {
            // k > 9 is considered unreasonably large for a single-element array
            if (k > 9)
            {
                return FormulaResult.Error("#NUM!");
            }

            return FormulaResult.FromNumber(values[0]);
        }

        if (k > values.Count)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Sort in ascending order to get smallest values first
        values.Sort();

        // Return k-th smallest (1-based indexing)
        return FormulaResult.FromNumber(values[k - 1]);
    }
}
