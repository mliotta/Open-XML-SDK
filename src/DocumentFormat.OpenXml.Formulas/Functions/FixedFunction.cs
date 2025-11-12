// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FIXED function.
/// FIXED(number, [decimals], [no_commas]) - formats number as text with fixed decimals.
/// </summary>
public sealed class FixedFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FixedFunction Instance = new();

    private FixedFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FIXED";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 3)
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

        // Get no_commas (default FALSE)
        bool noCommas = false;
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Boolean)
            {
                noCommas = args[2].BoolValue;
            }
            else if (args[2].Type == CellValueType.Number)
            {
                noCommas = args[2].NumericValue != 0;
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Round to specified decimals
        var rounded = System.Math.Round(number, decimals, MidpointRounding.AwayFromZero);

        // Format the number
        string result;
        if (noCommas)
        {
            // No commas - use fixed-point format
            result = rounded.ToString($"F{decimals}", CultureInfo.InvariantCulture);
        }
        else
        {
            // With commas - use number format
            result = rounded.ToString($"N{decimals}", CultureInfo.InvariantCulture);
        }

        return CellValue.FromString(result);
    }
}
