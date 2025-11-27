// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DVARP function.
/// DVARP(database, field, criteria) - Population variance of values in field that meet criteria.
/// Phase 0: Simplified implementation accepting individual values.
/// Future: Full range support with database headers and criteria ranges.
/// </summary>
public sealed class DVarPFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DVarPFunction Instance = new();

    private DVarPFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DVARP";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // DVARP requires exactly 3 arguments: database, field, criteria
        if (args.Length != 3)
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

        // Phase 0 simplified implementation:
        // For now, treat as single value operation
        // database = single value to check
        // field = ignored (would be column name/number in full implementation)
        // criteria = comparison criteria

        var database = args[0];
        var criteria = args[2];

        var values = new System.Collections.Generic.List<double>();

        // Apply criteria matching logic
        if (MatchesCriteria(database, criteria))
        {
            if (database.Type == FormulaResultType.Number)
            {
                values.Add(database.NumericValue);
            }
        }

        // DVARP requires at least 1 value
        if (values.Count == 0)
        {
            return FormulaResult.Error("#DIV/0!");
        }

        // For a single value, population variance is 0
        if (values.Count == 1)
        {
            return FormulaResult.FromNumber(0.0);
        }

        // Calculate population variance
        var mean = 0.0;
        foreach (var value in values)
        {
            mean += value;
        }

        mean /= values.Count;

        var sumSquaredDiff = 0.0;
        foreach (var value in values)
        {
            var diff = value - mean;
            sumSquaredDiff += diff * diff;
        }

        // Population variance (divide by n)
        var variance = sumSquaredDiff / values.Count;

        return FormulaResult.FromNumber(variance);
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
