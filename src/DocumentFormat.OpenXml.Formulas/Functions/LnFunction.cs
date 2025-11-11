// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LN function.
/// LN(number) - returns the natural logarithm (base e) of a number.
/// </summary>
public sealed class LnFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LnFunction Instance = new();

    private LnFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LN";

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

        if (number <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        var result = System.Math.Log(number);
        return CellValue.FromNumber(result);
    }
}
