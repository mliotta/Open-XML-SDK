// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BASE function.
/// BASE(number, radix, [min_length]) - converts a number into a text representation with the given radix (base).
/// Radix must be between 2 and 36.
/// Optional min_length parameter pads the result with leading zeros.
/// </summary>
public sealed class BaseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BaseFunction Instance = new();

    private BaseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BASE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
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

        var number = (long)args[0].NumericValue;

        if (number < 0)
        {
            return CellValue.Error("#NUM!");
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

        // Third argument: minimum length (optional)
        int minLength = 0;
        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            minLength = (int)args[2].NumericValue;

            if (minLength < 0 || minLength > 255)
            {
                return CellValue.Error("#NUM!");
            }
        }

        // Convert the number to the specified base
        string result = ConvertToBase(number, radix);

        // Pad with leading zeros if necessary
        if (result.Length < minLength)
        {
            result = result.PadLeft(minLength, '0');
        }

        // Check if result exceeds maximum length
        if (result.Length > 255)
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromString(result);
    }

    private static string ConvertToBase(long number, int radix)
    {
        if (number == 0)
        {
            return "0";
        }

        const string digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var result = new StringBuilder();

        while (number > 0)
        {
            result.Insert(0, digits[(int)(number % radix)]);
            number /= radix;
        }

        return result.ToString();
    }
}
