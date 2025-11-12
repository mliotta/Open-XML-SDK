// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DEC2OCT function.
/// DEC2OCT(number, [places]) - converts decimal to octal.
/// </summary>
public sealed class Dec2OctFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Dec2OctFunction Instance = new();

    private Dec2OctFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DEC2OCT";

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

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // Validate range: -536870912 to 536870911 (30-bit signed)
        if (number < -536870912.0 || number > 536870911.0)
        {
            return CellValue.Error("#NUM!");
        }

        // Truncate to integer
        var intValue = (int)System.Math.Floor(number);

        int places = 0;
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

            places = (int)System.Math.Floor(args[1].NumericValue);

            if (places < 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (places > 10)
            {
                return CellValue.Error("#NUM!");
            }
        }

        string octalString;

        // Handle negative numbers using two's complement
        if (intValue < 0)
        {
            // Convert to 30-bit two's complement
            long twosComplement = 0x40000000L + intValue; // 2^30
            octalString = Convert.ToString(twosComplement, 8);
        }
        else
        {
            octalString = Convert.ToString(intValue, 8);
        }

        // Apply padding if places specified
        if (places > 0)
        {
            if (octalString.Length > places)
            {
                return CellValue.Error("#NUM!");
            }

            octalString = octalString.PadLeft(places, '0');
        }

        return CellValue.FromString(octalString);
    }
}
