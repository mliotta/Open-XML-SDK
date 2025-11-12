// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACOTH function.
/// ACOTH(number) - returns the inverse hyperbolic cotangent of a number.
/// ACOTH(x) = 0.5 * LN((x+1)/(x-1))
/// Valid for |x| > 1
/// </summary>
public sealed class AcothFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AcothFunction Instance = new();

    private AcothFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACOTH";

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

        var x = args[0].NumericValue;

        // ACOTH is only defined for |x| > 1
        if (System.Math.Abs(x) <= 1.0)
        {
            return CellValue.Error("#NUM!");
        }

        // ACOTH(x) = 0.5 * ln((x+1)/(x-1))
        var result = 0.5 * System.Math.Log((x + 1) / (x - 1));

        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
