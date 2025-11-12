// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CSC function.
/// CSC(number) - returns the cosecant of an angle (number in radians).
/// CSC(x) = 1/SIN(x)
/// </summary>
public sealed class CscFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CscFunction Instance = new();

    private CscFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CSC";

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

        var sinValue = System.Math.Sin(args[0].NumericValue);

        // Check if sin is zero (would cause division by zero)
        if (System.Math.Abs(sinValue) < double.Epsilon)
        {
            return CellValue.Error("#DIV/0!");
        }

        var result = 1.0 / sinValue;

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
