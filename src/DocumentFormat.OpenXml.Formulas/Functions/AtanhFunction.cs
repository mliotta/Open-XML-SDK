// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ATANH function.
/// ATANH(number) - returns the inverse hyperbolic tangent of a number.
/// Number must be between -1 and 1 (exclusive).
/// </summary>
public sealed class AtanhFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AtanhFunction Instance = new();

    private AtanhFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ATANH";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
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

        if (number <= -1 || number >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // ATANH(x) = 0.5 * ln((1 + x) / (1 - x))
        var result = 0.5 * System.Math.Log((1 + number) / (1 - number));
        return CellValue.FromNumber(result);
    }
}
