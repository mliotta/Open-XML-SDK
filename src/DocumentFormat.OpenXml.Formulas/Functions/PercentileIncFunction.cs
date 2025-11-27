// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTILE.INC function.
/// PERCENTILE.INC(array, k) - returns the k-th percentile (0 &lt;= k &lt;= 1).
/// Uses linear interpolation between values (inclusive method).
/// This is the same as the legacy PERCENTILE function.
/// </summary>
public sealed class PercentileIncFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PercentileIncFunction Instance = new();

    private PercentileIncFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERCENTILE.INC";

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

        var k = args[1].NumericValue;

        // k must be between 0 and 1
        if (k < 0 || k > 1)
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

        // Sort values in ascending order
        values.Sort();

        // Calculate percentile using linear interpolation
        // Excel's PERCENTILE.INC uses (n-1) * k formula
        var n = values.Count;

        if (n == 1)
        {
            return FormulaResult.FromNumber(values[0]);
        }

        // Calculate position (0-based)
        var position = (n - 1) * k;
        var lowerIndex = (int)System.Math.Floor(position);
        var upperIndex = (int)System.Math.Ceiling(position);

        // If position is exact, return that value
        if (lowerIndex == upperIndex)
        {
            return FormulaResult.FromNumber(values[lowerIndex]);
        }

        // Linear interpolation between lower and upper values
        var lowerValue = values[lowerIndex];
        var upperValue = values[upperIndex];
        var fraction = position - lowerIndex;
        var result = lowerValue + fraction * (upperValue - lowerValue);

        return FormulaResult.FromNumber(result);
    }
}
