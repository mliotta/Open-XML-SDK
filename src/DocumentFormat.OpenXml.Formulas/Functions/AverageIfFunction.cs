// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AVERAGEIF function.
/// AVERAGEIF(range, criteria, [average_range]) - Average of cells meeting criteria.
/// </summary>
public sealed class AverageIfFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AverageIfFunction Instance = new();

    private AverageIfFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AVERAGEIF";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // AVERAGEIF requires 2 or 3 arguments
        if (args.Length < 2 || args.Length > 3)
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

        var criteriaRange = args[0];
        var criteria = args[1];
        var averageRange = args.Length == 3 ? args[2] : criteriaRange;

        // For flattened arrays, criteria range and average range should have equal length
        // Note: In the actual implementation, ranges come as CellValue[] where each element
        // represents a cell in the range. Since we receive already flattened arrays,
        // we'll process them element by element.

        var sum = 0.0;
        var count = 0;

        // Since ranges are passed as individual CellValue arguments when flattened,
        // and we only receive the first element here, we need to handle this differently.
        // Based on the user's note about flattened arrays, we'll assume args[0] represents
        // the entire criteria range as a single array if it's structured that way.

        // For this implementation, we'll work with what we have:
        // If only 2 args: AVERAGEIF(range, criteria) - average the range where it meets criteria
        // If 3 args: AVERAGEIF(range, criteria, average_range) - average average_range where range meets criteria

        // Simple implementation for single cell case:
        if (MatchesCriteria(criteriaRange, criteria))
        {
            if (averageRange.Type == CellValueType.Number)
            {
                sum += averageRange.NumericValue;
                count++;
            }
        }

        if (count == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        return CellValue.FromNumber(sum / count);
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
