// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOOKUP function.
/// LOOKUP(lookup_value, lookup_vector, [result_vector]) - vector/array lookup.
/// Find largest value &lt;= lookup_value in sorted vector.
/// </summary>
public sealed class LookupFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LookupFunction Instance = new();

    private LookupFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOOKUP";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Phase 0: Simplified implementation
        // For single value lookups, just return the lookup value if they match
        var lookupValue = args[0];
        var searchValue = args[1];

        // Check if values match (exact match for Phase 0)
        if (lookupValue.Type == CellValueType.Number && searchValue.Type == CellValueType.Number)
        {
            if (System.Math.Abs(lookupValue.NumericValue - searchValue.NumericValue) < 1e-10)
            {
                // If result vector provided, return it; otherwise return the search value
                if (args.Length == 3)
                {
                    return args[2];
                }
                return searchValue;
            }
        }
        else if (lookupValue.Type == CellValueType.Text && searchValue.Type == CellValueType.Text)
        {
            if (string.Equals(lookupValue.StringValue, searchValue.StringValue, StringComparison.OrdinalIgnoreCase))
            {
                if (args.Length == 3)
                {
                    return args[2];
                }
                return searchValue;
            }
        }

        // Full array/vector lookup requires array support
        // For Phase 0, return #N/A if no match found
        return CellValue.Error("#N/A");
    }
}
