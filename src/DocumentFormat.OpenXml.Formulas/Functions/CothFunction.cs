// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COTH function.
/// COTH(number) - returns the hyperbolic cotangent of a number.
/// COTH(x) = 1/TANH(x) = COSH(x)/SINH(x) = (e^x + e^(-x))/(e^x - e^(-x))
/// </summary>
public sealed class CothFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CothFunction Instance = new();

    private CothFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COTH";

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

        var tanhValue = System.Math.Tanh(args[0].NumericValue);

        // Check if tanh is zero (would cause division by zero)
        if (System.Math.Abs(tanhValue) < double.Epsilon)
        {
            return CellValue.Error("#DIV/0!");
        }

        var result = 1.0 / tanhValue;

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
