// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PROB function.
/// PROB(x_range, prob_range, [lower_limit], [upper_limit]) - returns the probability that values are between two limits.
/// </summary>
public sealed class ProbFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ProbFunction Instance = new();

    private ProbFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PROB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse arguments - structure is: x_values..., prob_values..., lower_limit, [upper_limit]
        // We need to determine the split point
        int totalArgs = args.Length;

        // Assume equal split for x and prob ranges, then lower/upper limits
        bool hasUpperLimit = totalArgs % 2 == 0;
        int limitCount = hasUpperLimit ? 2 : 1;
        int rangeSize = (totalArgs - limitCount) / 2;

        var xValues = new List<double>();
        var probValues = new List<double>();

        // Collect x values
        for (int i = 0; i < rangeSize; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                xValues.Add(args[i].NumericValue);
            }
        }

        // Collect prob values
        for (int i = rangeSize; i < rangeSize * 2; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                double prob = args[i].NumericValue;
                if (prob < 0 || prob > 1)
                {
                    return CellValue.Error("#NUM!");
                }
                probValues.Add(prob);
            }
        }

        if (xValues.Count != probValues.Count || xValues.Count == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Verify probabilities sum to 1 (with tolerance)
        double sumProb = 0;
        foreach (var p in probValues)
        {
            sumProb += p;
        }
        if (System.Math.Abs(sumProb - 1.0) > 0.0001)
        {
            return CellValue.Error("#NUM!");
        }

        // Get lower limit
        int lowerLimitIndex = rangeSize * 2;
        if (args[lowerLimitIndex].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double lowerLimit = args[lowerLimitIndex].NumericValue;

        // Get upper limit (or use lower limit if not provided)
        double upperLimit = lowerLimit;
        if (hasUpperLimit)
        {
            if (args[lowerLimitIndex + 1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }
            upperLimit = args[lowerLimitIndex + 1].NumericValue;
        }

        // Calculate probability sum for values in range
        double result = 0;
        for (int i = 0; i < xValues.Count; i++)
        {
            if (xValues[i] >= lowerLimit && xValues[i] <= upperLimit)
            {
                result += probValues[i];
            }
        }

        return CellValue.FromNumber(result);
    }
}
