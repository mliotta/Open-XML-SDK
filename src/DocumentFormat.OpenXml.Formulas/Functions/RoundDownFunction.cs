// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ROUNDDOWN function.
/// ROUNDDOWN(number, num_digits) - always rounds toward zero.
/// </summary>
public sealed class RoundDownFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RoundDownFunction Instance = new();

    private RoundDownFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ROUNDDOWN";

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

        var multiplier = System.Math.Pow(10, digits);
        var result = number >= 0
            ? System.Math.Floor(number * multiplier) / multiplier
            : System.Math.Ceiling(number * multiplier) / multiplier;

        return CellValue.FromNumber(result);
    }
}
