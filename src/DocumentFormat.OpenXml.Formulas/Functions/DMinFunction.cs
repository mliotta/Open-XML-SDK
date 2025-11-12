// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DMIN function.
/// DMIN(database, field, criteria) - Minimum value in field that meets criteria.
/// Phase 0: Simplified implementation accepting individual values.
/// Future: Full range support with database headers and criteria ranges.
/// </summary>
public sealed class DMinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DMinFunction Instance = new();

    private DMinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DMIN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // DMIN requires exactly 3 arguments: database, field, criteria
        if (args.Length != 3)
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

        // Phase 0 simplified implementation:
        // For now, treat as single value operation
        // database = single value to check
        // field = ignored (would be column name/number in full implementation)
        // criteria = comparison criteria

        var database = args[0];
        var criteria = args[2];

        var min = double.MaxValue;
        var hasValue = false;

        // Apply criteria matching logic
        if (MatchesCriteria(database, criteria))
        {
            if (database.Type == CellValueType.Number)
            {
                min = System.Math.Min(min, database.NumericValue);
                hasValue = true;
            }
        }

        if (!hasValue)
        {
            return CellValue.FromNumber(0);
        }

        return CellValue.FromNumber(min);
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
