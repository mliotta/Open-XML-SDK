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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument: text to convert
        if (args[0].IsError)
        {
            return args[0];
        }

        string text;
        if (args[0].Type == CellValueType.Text)
        {
            text = args[0].StringValue?.Trim() ?? string.Empty;
        }
        else if (args[0].Type == CellValueType.Number)
        {
            text = args[0].NumericValue.ToString("F0");
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        if (string.IsNullOrEmpty(text))
        {
            return CellValue.Error("#VALUE!");
        }

        // Second argument: radix (base)
        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var radix = (int)args[1].NumericValue;

        if (radix < 2 || radix > 36)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Convert from the specified base to decimal
            var result = Convert.ToInt64(text, radix);
            return CellValue.FromNumber(result);
        }
        catch (ArgumentException)
        {
            // Invalid characters for the specified base
            return CellValue.Error("#NUM!");
        }
        catch (FormatException)
        {
            return CellValue.Error("#NUM!");
        }
        catch (OverflowException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
