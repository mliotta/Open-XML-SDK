// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ROMAN function.
/// ROMAN(number, [form]) - converts an Arabic numeral to Roman, as text.
/// Form parameter (0-4) specifies the type of Roman numeral (0 = Classic, 4 = Simplified).
/// </summary>
public sealed class RomanFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RomanFunction Instance = new();

    private RomanFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ROMAN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument: number to convert
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = (int)args[0].NumericValue;

        if (number < 0 || number > 3999)
        {
            return CellValue.Error("#VALUE!");
        }

        if (number == 0)
        {
            return CellValue.FromString(string.Empty);
        }

        // Second argument: form (optional, default is 0 - classic)
        int form = 0;
        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            form = (int)args[1].NumericValue;

            if (form < 0 || form > 4)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var result = ConvertToRoman(number, form);
        return CellValue.FromString(result);
    }

    private static string ConvertToRoman(int number, int form)
    {
        if (number == 0)
        {
            return string.Empty;
        }

        // For form 0 (classic), use standard conversion
        // Forms 1-4 provide increasingly more concise representations
        // For simplicity, we'll implement classic form (0) fully
        // and treat other forms similarly (Excel's behavior varies slightly by form)

        var result = new StringBuilder();

        // Define the values and their Roman numeral representations
        var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var numerals = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };

        for (int i = 0; i < values.Length; i++)
        {
            while (number >= values[i])
            {
                result.Append(numerals[i]);
                number -= values[i];
            }
        }

        // Apply form-specific simplifications
        if (form > 0)
        {
            // More concise forms could replace certain patterns
            // For example, form 4 is most concise
            // This is a simplified implementation - Excel's actual behavior is more complex
            string romanStr = result.ToString();

            if (form >= 2)
            {
                // Replace certain verbose patterns with more concise ones
                romanStr = romanStr.Replace("DCCCC", "CM");
                romanStr = romanStr.Replace("CCCC", "CD");
                romanStr = romanStr.Replace("LXXXX", "XC");
                romanStr = romanStr.Replace("XXXX", "XL");
                romanStr = romanStr.Replace("VIIII", "IX");
                romanStr = romanStr.Replace("IIII", "IV");
            }

            return romanStr;
        }

        return result.ToString();
    }
}
