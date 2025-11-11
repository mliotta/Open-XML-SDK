// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTILE function.
/// PERCENTILE(array, k) - returns the k-th percentile (0 &lt;= k &lt;= 1).
/// Uses linear interpolation between values.
/// </summary>
public sealed class PercentileFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PercentileFunction Instance = new();

    private PercentileFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERCENTILE";

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

        var k = args[1].NumericValue;

        // k must be between 0 and 1
        if (k < 0 || k > 1)
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

        // Sort values in ascending order
        values.Sort();

        // Calculate percentile using linear interpolation
        // Excel's PERCENTILE uses (n-1) * k formula
        var n = values.Count;

        if (n == 1)
        {
            return CellValue.FromNumber(values[0]);
        }

        // Calculate position (0-based)
        var position = (n - 1) * k;
        var lowerIndex = (int)System.Math.Floor(position);
        var upperIndex = (int)System.Math.Ceiling(position);

        // If position is exact, return that value
        if (lowerIndex == upperIndex)
        {
            return CellValue.FromNumber(values[lowerIndex]);
        }

        // Linear interpolation between lower and upper values
        var lowerValue = values[lowerIndex];
        var upperValue = values[upperIndex];
        var fraction = position - lowerIndex;
        var result = lowerValue + fraction * (upperValue - lowerValue);

        return CellValue.FromNumber(result);
    }
}
