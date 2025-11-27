// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DECIMAL function.
/// DECIMAL(text, radix) - converts a text representation of a number in a given base into a decimal number.
/// Radix must be between 2 and 36.
/// </summary>
public sealed class DecimalFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DecimalFunction Instance = new();

    private DecimalFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DECIMAL";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // First argument: text to convert
        if (args[0].IsError)
        {
            return args[0];
        }

        string text;
        if (args[0].Type == FormulaResultType.Text)
        {
            text = args[0].StringValue?.Trim().ToUpperInvariant() ?? string.Empty;
        }
        else if (args[0].Type == FormulaResultType.Number)
        {
            text = args[0].NumericValue.ToString("F0");
        }
        else
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (string.IsNullOrEmpty(text))
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Second argument: radix (base)
        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var radix = (int)args[1].NumericValue;

        if (radix < 2 || radix > 36)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Convert from the specified base to decimal
            // Convert.ToInt64 only supports bases 2, 8, 10, 16, so we need a custom implementation
            long result = 0;
            foreach (char c in text)
            {
                int digit;
                if (c >= '0' && c <= '9')
                {
                    digit = c - '0';
                }
                else if (c >= 'A' && c <= 'Z')
                {
                    digit = c - 'A' + 10;
                }
                else
                {
                    // Invalid character
                    return FormulaResult.Error("#NUM!");
                }

                if (digit >= radix)
                {
                    // Digit is too large for the radix
                    return FormulaResult.Error("#NUM!");
                }

                result = result * radix + digit;
            }

            return FormulaResult.FromNumber(result);
        }
        catch (OverflowException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
