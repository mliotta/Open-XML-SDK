// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TEXT function.
/// TEXT(value, format_text) - converts number to formatted text.
/// </summary>
public sealed class TextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TextFunction Instance = new();

    private TextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TEXT";

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

        var value = args[0];
        var format = args[1].StringValue;

        // Convert value to number if possible
        double number;
        if (value.Type == CellValueType.Number)
        {
            number = value.NumericValue;
        }
        else if (value.Type == CellValueType.Text && double.TryParse(value.StringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
        {
            number = parsed;
        }
        else
        {
            return CellValue.FromString(value.StringValue);
        }

        // Basic format handling
        try
        {
            // Convert Excel format to .NET format (simplified)
            var dotNetFormat = format
                .Replace("#,##0", "N0")
                .Replace("0.00", "F2")
                .Replace("0", "F0");

            var result = number.ToString(dotNetFormat, CultureInfo.InvariantCulture);
            return CellValue.FromString(result);
        }
        catch
        {
            // If format fails, return plain number
            return CellValue.FromString(number.ToString(CultureInfo.InvariantCulture));
        }
    }
}
