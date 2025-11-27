// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COUNTIF function.
/// COUNTIF(range, criteria) - Counts cells that meet a criteria.
/// </summary>
public sealed class CountIfFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CountIfFunction Instance = new();

    private CountIfFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COUNTIF";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // COUNTIF requires exactly 2 arguments
        if (args.Length != 2)
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

        var criteriaRange = args[0];
        var criteria = args[1];

        var count = 0;

        // For single cell case:
        if (MatchesCriteria(criteriaRange, criteria))
        {
            count++;
        }

        return FormulaResult.FromNumber(count);
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
