// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DOLLAR function.
/// DOLLAR(number, [decimals]) - formats number as currency text.
/// </summary>
public sealed class DollarFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DollarFunction Instance = new();

    private DollarFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DOLLAR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        // Get number
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // Get decimals (default 2)
        int decimals = 2;
        if (args.Length >= 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            decimals = (int)args[1].NumericValue;
            if (decimals < 0)
            {
                decimals = 0;
            }
        }

        // Round to specified decimals
        var rounded = System.Math.Round(number, decimals, MidpointRounding.AwayFromZero);

        // Format with dollar sign and commas
        var result = "$" + rounded.ToString($"N{decimals}", CultureInfo.InvariantCulture);

        return CellValue.FromString(result);
    }
}
