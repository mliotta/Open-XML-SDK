// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRUNC function.
/// TRUNC(number, [num_digits]) - truncates a number to a specified precision.
/// </summary>
public sealed class TruncFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TruncFunction Instance = new();

    private TruncFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRUNC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var numDigits = 0;

        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            numDigits = (int)args[1].NumericValue;
        }

        // Truncate by multiplying, truncating, and dividing back
        var multiplier = System.Math.Pow(10, numDigits);
        var result = System.Math.Truncate(number * multiplier) / multiplier;
        return CellValue.FromNumber(result);
    }
}
