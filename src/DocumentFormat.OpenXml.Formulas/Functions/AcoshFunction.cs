// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACOSH function.
/// ACOSH(number) - returns the inverse hyperbolic cosine of a number.
/// Number must be greater than or equal to 1.
/// </summary>
public sealed class AcoshFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AcoshFunction Instance = new();

    private AcoshFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACOSH";

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

        if (number < 1)
        {
            return CellValue.Error("#NUM!");
        }

        // ACOSH(x) = ln(x + sqrt(x^2 - 1))
        var result = System.Math.Log(number + System.Math.Sqrt(number * number - 1));
        return CellValue.FromNumber(result);
    }
}
