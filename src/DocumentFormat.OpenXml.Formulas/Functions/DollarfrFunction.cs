// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DOLLARFR function.
/// DOLLARFR(decimal_dollar, fraction) - converts a dollar price expressed as a decimal into a fractional number.
/// </summary>
public sealed class DollarfrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DollarfrFunction Instance = new();

    private DollarfrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DOLLARFR";

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

        var decimalDollar = args[0].NumericValue;
        var fraction = (int)args[1].NumericValue;

        // Fraction must be positive
        if (fraction <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Split into integer and decimal parts
            var integerPart = System.Math.Floor(System.Math.Abs(decimalDollar));
            var decimalPart = System.Math.Abs(decimalDollar) - integerPart;

            // Convert decimal part to fractional representation
            // Example: 1.125 with fraction 16 means (0.125 * 16) = 2, so result is 1.02
            var numerator = System.Math.Round(decimalPart * fraction);
            var fractionalPart = numerator / 100.0;

            // Reconstruct with proper sign
            var result = decimalDollar < 0 ? -(integerPart + fractionalPart) : (integerPart + fractionalPart);

            return FormulaResult.FromNumber(result);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
