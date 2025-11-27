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

        var fractionalDollar = args[0].NumericValue;
        var fraction = (int)args[1].NumericValue;

        // Fraction must be positive
        if (fraction <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Split into integer and fractional parts
            var integerPart = System.Math.Floor(System.Math.Abs(fractionalDollar));
            var fractionalPart = System.Math.Abs(fractionalDollar) - integerPart;

            // The fractional part represents the numerator when interpreted as integer digits
            // Example: 1.02 with fraction 16 means 1 + 2/16 = 1.125
            // Example: 1.1 with fraction 32 means 1 + 1/32 = 1.03125
            // Parse the fractional digits as an integer
            var fracStr = fractionalPart.ToString("F10", System.Globalization.CultureInfo.InvariantCulture);
            var decimalIndex = fracStr.IndexOf('.');
            var digitsStr = decimalIndex >= 0 ? fracStr.Substring(decimalIndex + 1).TrimEnd('0') : "0";
            var numerator = string.IsNullOrEmpty(digitsStr) ? 0 : int.Parse(digitsStr);
            var decimalPart = (double)numerator / fraction;

            // Reconstruct with proper sign
            var result = fractionalDollar < 0 ? -(integerPart + decimalPart) : (integerPart + decimalPart);

            return FormulaResult.FromNumber(result);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
