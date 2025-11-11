// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SQRT function.
/// SQRT(number) - returns the square root of a number.
/// </summary>
public sealed class SqrtFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SqrtFunction Instance = new();

    private SqrtFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SQRT";

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

        if (number < 0)
        {
            return CellValue.Error("#NUM!");
        }

        var result = System.Math.Sqrt(number);
        return CellValue.FromNumber(result);
    }
}
