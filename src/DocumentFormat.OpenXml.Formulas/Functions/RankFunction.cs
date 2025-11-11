// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RANK function.
/// RANK(number, ref, [order]) - returns rank of number in list.
/// </summary>
public sealed class RankFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RankFunction Instance = new();

    private RankFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RANK";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // Collect all numeric values from the reference (args[1] onwards)
        // Note: Due to how the compiler flattens arguments, we cannot reliably distinguish
        // between range values and an optional order parameter. For now, we treat all
        // args[1..] as the range and default to descending order (order=0).
        var values = new List<double>();

        for (var i = 1; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                values.Add(args[i].NumericValue);
            }
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Check if number exists in the list
        if (!values.Contains(number))
        {
            return CellValue.Error("#N/A");
        }

        // Calculate rank in descending order (largest = rank 1)
        var rank = values.Count(v => v > number) + 1;

        return CellValue.FromNumber(rank);
    }
}
