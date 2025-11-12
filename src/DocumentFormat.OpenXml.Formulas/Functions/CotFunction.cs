// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COT function.
/// COT(number) - returns the cotangent of an angle (number in radians).
/// COT(x) = 1/TAN(x) = COS(x)/SIN(x)
/// </summary>
public sealed class CotFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CotFunction Instance = new();

    private CotFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COT";

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

        var tanValue = System.Math.Tan(args[0].NumericValue);

        // Check if tan is zero (would cause division by zero)
        if (System.Math.Abs(tanValue) < double.Epsilon)
        {
            return CellValue.Error("#DIV/0!");
        }

        var result = 1.0 / tanValue;

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
