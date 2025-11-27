// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ARRAYTOTEXT function.
/// ARRAYTOTEXT(array, [format]) - converts array to text representation.
/// </summary>
public sealed class ArrayToTextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ArrayToTextFunction Instance = new();

    private ArrayToTextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ARRAYTOTEXT";

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

        // For single values (not arrays), convert similar to VALUETOTEXT
        var result = FormatValue(value, format);
        return FormulaResult.FromString(result);
    }

    private static string FormatValue(FormulaResult value, int format)
    {
        switch (value.Type)
        {
            case FormulaResultType.Text:
                return format == 1
                    ? $"\"{value.StringValue}\""
                    : value.StringValue;

            case FormulaResultType.Number:
                return value.NumericValue.ToString(CultureInfo.InvariantCulture);

            case FormulaResultType.Boolean:
                return value.BoolValue ? "TRUE" : "FALSE";

            case FormulaResultType.Empty:
                return string.Empty;

            case FormulaResultType.Error:
                return value.ErrorValue ?? "#VALUE!";

            default:
                return value.StringValue;
        }
    }
}
