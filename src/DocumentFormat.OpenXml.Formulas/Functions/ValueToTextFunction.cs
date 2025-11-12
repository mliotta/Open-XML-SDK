// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VALUETOTEXT function.
/// VALUETOTEXT(value, [format]) - converts value to text in specified format.
/// </summary>
public sealed class ValueToTextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ValueToTextFunction Instance = new();

    private ValueToTextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VALUETOTEXT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors
        if (args[0].IsError)
        {
            return args[0];
        }

        // Default format: 0 = concise, 1 = strict
        var format = 0;

        if (args.Length >= 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type == CellValueType.Number)
            {
                format = (int)args[1].NumericValue;
                if (format != 0 && format != 1)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        var value = args[0];

        // Convert value to text based on type
        switch (value.Type)
        {
            case CellValueType.Text:
                return format == 1
                    ? CellValue.FromString($"\"{value.StringValue}\"")
                    : CellValue.FromString(value.StringValue);

            case CellValueType.Number:
                return CellValue.FromString(value.NumericValue.ToString(CultureInfo.InvariantCulture));

            case CellValueType.Boolean:
                return CellValue.FromString(value.BoolValue ? "TRUE" : "FALSE");

            case CellValueType.Empty:
                return CellValue.FromString(string.Empty);

            case CellValueType.Error:
                return CellValue.FromString(value.ErrorValue ?? "#VALUE!");

            default:
                return CellValue.FromString(value.StringValue);
        }
    }
}
