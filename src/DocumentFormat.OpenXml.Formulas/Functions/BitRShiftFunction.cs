// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BITRSHIFT function.
/// BITRSHIFT(number, shift_amount) - returns a number shifted right by shift_amount bits.
/// </summary>
public sealed class BitRShiftFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BitRShiftFunction Instance = new();

    private BitRShiftFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BITRSHIFT";

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

        var number = (long)args[0].NumericValue;
        var shift = (int)args[1].NumericValue;

        // Must be non-negative and fit in 48 bits (Excel's limit)
        if (number < 0 || number > 281474976710655)
        {
            return CellValue.Error("#NUM!");
        }

        // Shift amount must be reasonable
        if (shift < -53 || shift > 53)
        {
            return CellValue.Error("#NUM!");
        }

        long result;
        if (shift >= 0)
        {
            result = number >> shift;
        }
        else
        {
            result = number << (-shift);
        }

        // Result must fit in 48 bits
        if (result < 0 || result > 281474976710655)
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
