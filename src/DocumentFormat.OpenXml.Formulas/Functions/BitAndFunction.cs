// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BITAND function.
/// BITAND(number1, number2) - returns a bitwise AND of two numbers.
/// </summary>
public sealed class BitAndFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BitAndFunction Instance = new();

    private BitAndFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BITAND";

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

        var num1 = (long)args[0].NumericValue;
        var num2 = (long)args[1].NumericValue;

        // Must be non-negative and fit in 48 bits (Excel's limit)
        if (num1 < 0 || num2 < 0 || num1 > 281474976710655 || num2 > 281474976710655)
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(num1 & num2);
    }
}
