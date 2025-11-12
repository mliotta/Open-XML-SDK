// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ASINH function.
/// ASINH(number) - returns the inverse hyperbolic sine of a number.
/// </summary>
public sealed class AsinhFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AsinhFunction Instance = new();

    private AsinhFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ASINH";

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

        var value = args[0].NumericValue;

        // ASINH(x) = ln(x + sqrt(x^2 + 1))
        var result = System.Math.Log(value + System.Math.Sqrt(value * value + 1));
        return CellValue.FromNumber(result);
    }
}
