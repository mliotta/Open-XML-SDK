// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FLOOR function.
/// FLOOR(number, significance) - rounds down to the nearest multiple of significance.
/// </summary>
public sealed class FloorFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FloorFunction Instance = new();

    private FloorFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FLOOR";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var significance = args[1].NumericValue;

        if (significance == 0)
        {
            return FormulaResult.FromNumber(0);
        }

        // Excel FLOOR behavior: if number and significance have different signs, return #NUM!
        if ((number > 0 && significance < 0) || (number < 0 && significance > 0))
        {
            return FormulaResult.Error("#NUM!");
        }

        // Round down to nearest multiple of significance
        var result = System.Math.Floor(number / significance) * significance;
        return FormulaResult.FromNumber(result);
    }
}
