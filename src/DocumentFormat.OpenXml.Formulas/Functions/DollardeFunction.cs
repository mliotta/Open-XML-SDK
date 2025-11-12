// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DOLLARDE function.
/// DOLLARDE(fractional_dollar, fraction) - converts a dollar price expressed as a fraction into a decimal number.
/// </summary>
public sealed class DollardeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DollardeFunction Instance = new();

    private DollardeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DOLLARDE";

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

        var fractionalDollar = args[0].NumericValue;
        var fraction = (int)args[1].NumericValue;

        // Fraction must be positive
        if (fraction <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Split into integer and fractional parts
            var integerPart = System.Math.Floor(System.Math.Abs(fractionalDollar));
            var fractionalPart = System.Math.Abs(fractionalDollar) - integerPart;

            // Convert fractional part from base fraction to decimal
            var decimalPart = fractionalPart / fraction;

            // Reconstruct with proper sign
            var result = fractionalDollar < 0 ? -(integerPart + decimalPart) : (integerPart + decimalPart);

            return CellValue.FromNumber(result);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
