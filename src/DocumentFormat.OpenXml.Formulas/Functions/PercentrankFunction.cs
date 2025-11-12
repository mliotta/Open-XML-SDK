// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTRANK function (legacy, same as PERCENTRANK.INC).
/// PERCENTRANK(array, x, [significance]) - returns the rank of a value as a percentage of the dataset.
/// </summary>
public sealed class PercentrankFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PercentrankFunction Instance = new();

    private PercentrankFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERCENTRANK";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get the value to find rank for (last or second-to-last argument)
        double x;
        int significance = 3; // default

        // Determine if we have a significance argument
        var xIndex = args.Length - 1;
        bool hasSignificance = false;

        // Try to parse last argument as significance
        if (args.Length >= 3 && args[args.Length - 1].Type == CellValueType.Number)
        {
            var lastVal = args[args.Length - 1].NumericValue;
            if (lastVal >= 1 && lastVal <= 15)
            {
                significance = (int)lastVal;
                hasSignificance = true;
                xIndex = args.Length - 2;
            }
        }

        if (args[xIndex].IsError)
        {
            return args[xIndex];
        }

        if (args[xIndex].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        x = args[xIndex].NumericValue;

        // Collect array values (all arguments except x and optional significance)
        var values = new List<double>();
        int endIndex = hasSignificance ? args.Length - 2 : args.Length - 1;

        for (int i = 0; i < endIndex; i++)
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

        // Sort values
        var sorted = values.OrderBy(v => v).ToArray();

        // Check if x is within range
        if (x < sorted[0] || x > sorted[sorted.Length - 1])
        {
            return CellValue.Error("#N/A");
        }

        // Find exact match or interpolate
        double rank;
        int exactIndex = System.Array.IndexOf(sorted, x);

        if (exactIndex >= 0)
        {
            // Exact match
            rank = (double)exactIndex / (sorted.Length - 1);
        }
        else
        {
            // Interpolate
            int upperIndex = 0;
            for (int i = 0; i < sorted.Length; i++)
            {
                if (sorted[i] > x)
                {
                    upperIndex = i;
                    break;
                }
            }

            int lowerIndex = upperIndex - 1;
            double lowerValue = sorted[lowerIndex];
            double upperValue = sorted[upperIndex];

            double lowerRank = (double)lowerIndex / (sorted.Length - 1);
            double upperRank = (double)upperIndex / (sorted.Length - 1);

            // Linear interpolation
            rank = lowerRank + (upperRank - lowerRank) * (x - lowerValue) / (upperValue - lowerValue);
        }

        // Round to specified significance
        double multiplier = System.Math.Pow(10, significance);
        double result = System.Math.Round(rank * multiplier) / multiplier;

        return CellValue.FromNumber(result);
    }
}
