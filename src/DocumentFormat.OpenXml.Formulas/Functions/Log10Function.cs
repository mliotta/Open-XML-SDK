// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOG10 function.
/// LOG10(number) - returns the base-10 logarithm of a number.
/// </summary>
public sealed class Log10Function : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Log10Function Instance = new();

    private Log10Function()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOG10";

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

        var result = System.Math.Log10(number);
        return CellValue.FromNumber(result);
    }
}
