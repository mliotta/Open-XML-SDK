// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RANK.AVG function.
/// RANK.AVG(number, ref, [order]) - returns average rank of number when there are duplicates.
/// </summary>
public sealed class RankAvgFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RankAvgFunction Instance = new();

    private RankAvgFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RANK.AVG";

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

        // Collect all numeric values from the reference
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

        // Calculate average rank for duplicates (descending order)
        // Count how many values are greater than number
        var greaterCount = values.Count(v => v > number);

        // Count how many values equal number
        var equalCount = values.Count(v => v == number);

        // Average rank is: (first_rank + last_rank) / 2
        // first_rank = greaterCount + 1
        // last_rank = greaterCount + equalCount
        double avgRank = (greaterCount + 1 + greaterCount + equalCount) / 2.0;

        return CellValue.FromNumber(avgRank);
    }
}
