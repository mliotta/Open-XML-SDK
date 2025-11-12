// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BITOR function.
/// BITOR(number1, number2) - returns a bitwise OR of two numbers.
/// </summary>
public sealed class BitOrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BitOrFunction Instance = new();

    private BitOrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BITOR";

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

        return CellValue.FromNumber(num1 | num2);
    }
}
