// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the UNICHAR function.
/// UNICHAR(number) - returns Unicode character for code point.
/// </summary>
public sealed class UnicharFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly UnicharFunction Instance = new();

    private UnicharFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "UNICHAR";

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

        var number = (int)args[0].NumericValue;

        // Valid Unicode code points are 1-1114111 (0x10FFFF) in Excel
        // Excluding surrogates range 0xD800-0xDFFF
        if (number < 1 || number > 1114111)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for surrogate range
        if (number >= 0xD800 && number <= 0xDFFF)
        {
            return CellValue.Error("#VALUE!");
        }

        string character;
        if (number <= 0xFFFF)
        {
            // Basic Multilingual Plane - single char
            character = ((char)number).ToString();
        }
        else
        {
            // Supplementary planes - requires surrogate pair
            character = char.ConvertFromUtf32(number);
        }

        return CellValue.FromString(character);
    }
}
