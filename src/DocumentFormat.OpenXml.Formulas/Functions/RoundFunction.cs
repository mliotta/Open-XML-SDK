// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ROUND function.
/// ROUND(number, num_digits) - rounds to specified digits using Excel's rounding mode (round half away from zero).
/// </summary>
public sealed class RoundFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RoundFunction Instance = new();

    private RoundFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ROUND";

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

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var digits = (int)args[1].NumericValue;

        var result = System.Math.Round(number, digits, MidpointRounding.AwayFromZero);
        return CellValue.FromNumber(result);
    }
}
