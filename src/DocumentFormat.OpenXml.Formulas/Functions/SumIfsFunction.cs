// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUMIFS function.
/// SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) - Sums cells that meet multiple criteria.
/// </summary>
public sealed class SumIfsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumIfsFunction Instance = new();

    private SumIfsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUMIFS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // SUMIFS requires at least 3 arguments (sum_range, criteria_range1, criteria1)
        // and additional criteria pairs (criteria_range, criteria)
        // Total arguments must be odd (sum_range + pairs of criteria_range/criteria)
        if (args.Length < 3 || args.Length % 2 == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }
        }

        var sumRange = args[0];

        // Check all criteria (AND logic)
        bool allCriteriaMet = true;
        for (int i = 1; i < args.Length; i += 2)
        {
            var criteriaRange = args[i];
            var criteria = args[i + 1];

            if (!MatchesCriteria(criteriaRange, criteria))
            {
                allCriteriaMet = false;
                break;
            }
        }

        var sum = 0.0;

        // If all criteria are met, include in sum
        if (allCriteriaMet)
        {
            if (sumRange.Type == CellValueType.Number)
            {
                sum += sumRange.NumericValue;
            }
        }

        return CellValue.FromNumber(sum);
    }

    private static bool MatchesCriteria(CellValue value, CellValue criteria)
    {
        // Handle criteria as a comparison operator + value
        if (criteria.Type == CellValueType.Text)
        {
            var criteriaText = criteria.StringValue;

            // Check for operators: >, <, >=, <=, <>, =
            if (criteriaText.StartsWith(">="))
            {
                if (double.TryParse(criteriaText.Substring(2), out var threshold))
                {
                    return value.Type == CellValueType.Number && value.NumericValue >= threshold;
                }
            }
            else if (criteriaText.StartsWith("<="))
            {
                if (double.TryParse(criteriaText.Substring(2), out var threshold))
                {
                    return value.Type == CellValueType.Number && value.NumericValue <= threshold;
                }
            }
            else if (criteriaText.StartsWith("<>"))
            {
                var compareValue = criteriaText.Substring(2);
                if (double.TryParse(compareValue, out var numValue))
                {
                    return value.Type != CellValueType.Number || value.NumericValue != numValue;
                }
                else
                {
                    return value.Type != CellValueType.Text || !value.StringValue.Equals(compareValue, StringComparison.OrdinalIgnoreCase);
                }
            }
            else if (criteriaText.StartsWith(">"))
            {
                if (double.TryParse(criteriaText.Substring(1), out var threshold))
                {
                    return value.Type == CellValueType.Number && value.NumericValue > threshold;
                }
            }
            else if (criteriaText.StartsWith("<"))
            {
                if (double.TryParse(criteriaText.Substring(1), out var threshold))
                {
                    return value.Type == CellValueType.Number && value.NumericValue < threshold;
                }
            }
            else if (criteriaText.StartsWith("="))
            {
                var compareValue = criteriaText.Substring(1);
                if (double.TryParse(compareValue, out var numValue))
                {
                    return value.Type == CellValueType.Number && value.NumericValue == numValue;
                }
                else
                {
                    return value.Type == CellValueType.Text && value.StringValue.Equals(compareValue, StringComparison.OrdinalIgnoreCase);
                }
            }
            else
            {
                // Direct text comparison (case-insensitive)
                return value.Type == CellValueType.Text && value.StringValue.Equals(criteriaText, StringComparison.OrdinalIgnoreCase);
            }
        }
        else if (criteria.Type == CellValueType.Number)
        {
            // Direct numeric comparison
            return value.Type == CellValueType.Number && value.NumericValue == criteria.NumericValue;
        }
        else if (criteria.Type == CellValueType.Boolean)
        {
            // Boolean comparison
            return value.Type == CellValueType.Boolean && value.BoolValue == criteria.BoolValue;
        }

        return false;
    }
}
