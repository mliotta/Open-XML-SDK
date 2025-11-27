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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return FormulaResult.Error("#VALUE!");
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

            if (args[1].Type == FormulaResultType.Number)
            {
                format = (int)args[1].NumericValue;
                if (format != 0 && format != 1)
                {
                    return FormulaResult.Error("#VALUE!");
                }
            }
        }

        var value = args[0];

        // Convert value to text based on type
        switch (value.Type)
        {
            case FormulaResultType.Text:
                return format == 1
                    ? FormulaResult.FromString($"\"{value.StringValue}\"")
                    : FormulaResult.FromString(value.StringValue);

            case FormulaResultType.Number:
                return FormulaResult.FromString(value.NumericValue.ToString(CultureInfo.InvariantCulture));

            case FormulaResultType.Boolean:
                return FormulaResult.FromString(value.BoolValue ? "TRUE" : "FALSE");

            case FormulaResultType.Empty:
                return FormulaResult.FromString(string.Empty);

            case FormulaResultType.Error:
                return FormulaResult.FromString(value.ErrorValue ?? "#VALUE!");

            default:
                return FormulaResult.FromString(value.StringValue);
        }
    }
}
