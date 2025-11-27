// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MINIFS function.
/// MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...) - Minimum of cells meeting multiple criteria.
/// </summary>
public sealed class MinIfsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MinIfsFunction Instance = new();

    private MinIfsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MINIFS";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // MINIFS requires at least 3 arguments (min_range, criteria_range1, criteria1)
        // and additional criteria pairs (criteria_range, criteria)
        // Total arguments must be odd (min_range + pairs of criteria_range/criteria)
        if (args.Length < 3 || args.Length % 2 == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in arguments
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }
        }

        var minRange = args[0];

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

        var min = double.MaxValue;
        var hasValue = false;

        // If all criteria are met, include in min calculation
        if (allCriteriaMet)
        {
            if (minRange.Type == FormulaResultType.Number)
            {
                min = System.Math.Min(min, minRange.NumericValue);
                hasValue = true;
            }
        }

        if (!hasValue)
        {
            return FormulaResult.FromNumber(0);
        }

        return FormulaResult.FromNumber(min);
    }

    private static bool MatchesCriteria(FormulaResult value, FormulaResult criteria)
    {
        // Handle criteria as a comparison operator + value
        if (criteria.Type == FormulaResultType.Text)
        {
            var criteriaText = criteria.StringValue;

            // Check for operators: >, <, >=, <=, <>, =
            if (criteriaText.StartsWith(">="))
            {
                if (double.TryParse(criteriaText.Substring(2), out var threshold))
                {
                    return value.Type == FormulaResultType.Number && value.NumericValue >= threshold;
                }
            }
            else if (criteriaText.StartsWith("<="))
            {
                if (double.TryParse(criteriaText.Substring(2), out var threshold))
                {
                    return value.Type == FormulaResultType.Number && value.NumericValue <= threshold;
                }
            }
            else if (criteriaText.StartsWith("<>"))
            {
                var compareValue = criteriaText.Substring(2);
                if (double.TryParse(compareValue, out var numValue))
                {
                    return value.Type != FormulaResultType.Number || value.NumericValue != numValue;
                }
                else
                {
                    return value.Type != FormulaResultType.Text || !value.StringValue.Equals(compareValue, StringComparison.OrdinalIgnoreCase);
                }
            }
            else if (criteriaText.StartsWith(">"))
            {
                if (double.TryParse(criteriaText.Substring(1), out var threshold))
                {
                    return value.Type == FormulaResultType.Number && value.NumericValue > threshold;
                }
            }
            else if (criteriaText.StartsWith("<"))
            {
                if (double.TryParse(criteriaText.Substring(1), out var threshold))
                {
                    return value.Type == FormulaResultType.Number && value.NumericValue < threshold;
                }
            }
            else if (criteriaText.StartsWith("="))
            {
                var compareValue = criteriaText.Substring(1);
                if (double.TryParse(compareValue, out var numValue))
                {
                    return value.Type == FormulaResultType.Number && value.NumericValue == numValue;
                }
                else
                {
                    return value.Type == FormulaResultType.Text && value.StringValue.Equals(compareValue, StringComparison.OrdinalIgnoreCase);
                }
            }
            else
            {
                // Direct text comparison (case-insensitive)
                return value.Type == FormulaResultType.Text && value.StringValue.Equals(criteriaText, StringComparison.OrdinalIgnoreCase);
            }
        }
        else if (criteria.Type == FormulaResultType.Number)
        {
            // Direct numeric comparison
            return value.Type == FormulaResultType.Number && value.NumericValue == criteria.NumericValue;
        }
        else if (criteria.Type == FormulaResultType.Boolean)
        {
            // Boolean comparison
            return value.Type == FormulaResultType.Boolean && value.BoolValue == criteria.BoolValue;
        }

        return false;
    }
}
