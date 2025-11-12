// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTRANK.EXC function.
/// PERCENTRANK.EXC(array, x, [significance]) - returns the rank of a value as a percentage (0 to 1 exclusive).
/// </summary>
public sealed class PercentrankExcFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PercentrankExcFunction Instance = new();

    private PercentrankExcFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERCENTRANK.EXC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get the value to find rank for
        double x;
        int significance = 3; // default

        var xIndex = args.Length - 1;
        bool hasSignificance = false;

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

        // Collect array values
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

        if (values.Count <= 1)
        {
            return CellValue.Error("#N/A");
        }

        // Sort values
        var sorted = values.OrderBy(v => v).ToArray();

        // Check if x is within range (exclusive)
        if (x <= sorted[0] || x >= sorted[sorted.Length - 1])
        {
            return CellValue.Error("#N/A");
        }

        // Find position and interpolate
        double rank;
        int exactIndex = System.Array.IndexOf(sorted, x);

        if (exactIndex >= 0)
        {
            // Exact match - use position formula for EXC
            rank = (double)(exactIndex + 1) / (sorted.Length + 1);
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

            double lowerRank = (double)(lowerIndex + 1) / (sorted.Length + 1);
            double upperRank = (double)(upperIndex + 1) / (sorted.Length + 1);

            rank = lowerRank + (upperRank - lowerRank) * (x - lowerValue) / (upperValue - lowerValue);
        }

        // Round to specified significance
        double multiplier = System.Math.Pow(10, significance);
        double result = System.Math.Round(rank * multiplier) / multiplier;

        return CellValue.FromNumber(result);
    }
}
