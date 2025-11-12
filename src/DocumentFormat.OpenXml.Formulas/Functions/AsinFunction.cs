// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ASIN function.
/// ASIN(number) - returns the arcsine (inverse sine) of a number in radians.
/// Number must be between -1 and 1.
/// </summary>
public sealed class AsinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AsinFunction Instance = new();

    private AsinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ASIN";

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

        if (number < -1 || number > 1)
        {
            return CellValue.Error("#NUM!");
        }

        var result = System.Math.Asin(number);
        return CellValue.FromNumber(result);
    }
}
