// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MAXA function.
/// MAXA(value1, [value2], ...) - Maximum value including text and logical values.
/// Text evaluates as 0, TRUE as 1, FALSE as 0, empty values are ignored.
/// </summary>
public sealed class MaxAFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MaxAFunction Instance = new();

    private MaxAFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MAXA";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        var max = double.MinValue;
        var hasValue = false;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == FormulaResultType.Number)
            {
                max = System.Math.Max(max, arg.NumericValue);
                hasValue = true;
            }
            else if (arg.Type == FormulaResultType.Boolean)
            {
                max = System.Math.Max(max, arg.BoolValue ? 1.0 : 0.0);
                hasValue = true;
            }
            else if (arg.Type == FormulaResultType.Text)
            {
                // Text values count as 0
                max = System.Math.Max(max, 0.0);
                hasValue = true;
            }
            // Empty values are ignored
        }

        if (!hasValue)
        {
            return FormulaResult.FromNumber(0);
        }

        return FormulaResult.FromNumber(max);
    }
}
