// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRIMMEAN function.
/// TRIMMEAN(array, percent) - returns the mean of the interior of a data set, excluding a percentage of outliers from the top and bottom.
/// </summary>
public sealed class TrimmeanFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrimmeanFunction Instance = new();

    private TrimmeanFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRIMMEAN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var values = new List<double>();

        // Extract numeric values from array
        if (args[0].Type == CellValueType.Number)
        {
            values.Add(args[0].NumericValue);
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get percent parameter
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        double percent = args[1].NumericValue;

        if (percent < 0 || percent >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Sort the values
        values.Sort();

        // Calculate number of values to trim from each end
        int n = values.Count;
        int trimCount = (int)System.Math.Floor(n * percent / 2.0);

        // If we would trim all values, return error
        if (trimCount * 2 >= n)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate mean of remaining values
        double sum = 0;
        int count = 0;
        for (int i = trimCount; i < n - trimCount; i++)
        {
            sum += values[i];
            count++;
        }

        if (count == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        double mean = sum / count;
        return CellValue.FromNumber(mean);
    }
}
