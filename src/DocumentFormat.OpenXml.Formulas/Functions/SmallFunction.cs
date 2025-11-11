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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
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
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var k = (int)args[1].NumericValue;

        if (k < 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Collect all numeric values
        var values = new List<double>();

        if (args[0].Type == CellValueType.Number)
        {
            values.Add(args[0].NumericValue);
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#NUM!");
        }

        if (k > values.Count)
        {
            return CellValue.Error("#NUM!");
        }

        // Sort in ascending order to get smallest values first
        values.Sort();

        // Return k-th smallest (1-based indexing)
        return CellValue.FromNumber(values[k - 1]);
    }
}
