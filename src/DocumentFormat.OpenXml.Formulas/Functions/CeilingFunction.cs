// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CEILING function.
/// CEILING(number, significance) - rounds up to the nearest multiple of significance.
/// </summary>
public sealed class CeilingFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CeilingFunction Instance = new();

    private CeilingFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CEILING";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var significance = args[1].NumericValue;

        if (significance == 0)
        {
            return CellValue.FromNumber(0);
        }

        // Excel CEILING behavior: if number and significance have different signs, return #NUM!
        if ((number > 0 && significance < 0) || (number < 0 && significance > 0))
        {
            return CellValue.Error("#NUM!");
        }

        // Round up to nearest multiple of significance
        var result = System.Math.Ceiling(number / significance) * significance;
        return CellValue.FromNumber(result);
    }
}
