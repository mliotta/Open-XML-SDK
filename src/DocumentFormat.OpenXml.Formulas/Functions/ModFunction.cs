// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MOD function.
/// MOD(number, divisor) - returns the remainder after division.
/// </summary>
public sealed class ModFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ModFunction Instance = new();

    private ModFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MOD";

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
        var divisor = args[1].NumericValue;

        if (divisor == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Excel MOD uses: MOD(n, d) = n - d*INT(n/d)
        // This matches Excel's behavior for negative numbers
        var result = number - divisor * System.Math.Floor(number / divisor);
        return CellValue.FromNumber(result);
    }
}
