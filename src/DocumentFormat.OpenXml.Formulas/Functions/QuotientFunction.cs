// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the QUOTIENT function.
/// QUOTIENT(numerator, denominator) - returns the integer portion of a division.
/// Truncates toward zero (same as TRUNC).
/// </summary>
public sealed class QuotientFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly QuotientFunction Instance = new();

    private QuotientFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "QUOTIENT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
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

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var numerator = args[0].NumericValue;
        var denominator = args[1].NumericValue;

        // Denominator cannot be zero
        if (denominator == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate division and truncate toward zero
        var result = System.Math.Truncate(numerator / denominator);

        return CellValue.FromNumber(result);
    }
}
