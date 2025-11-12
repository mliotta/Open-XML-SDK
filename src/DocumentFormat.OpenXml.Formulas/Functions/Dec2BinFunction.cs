// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DEC2BIN function.
/// DEC2BIN(number, [places]) - converts decimal to binary.
/// </summary>
public sealed class Dec2BinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Dec2BinFunction Instance = new();

    private Dec2BinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DEC2BIN";

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

        // Validate range: -512 to 511 (10-bit signed)
        if (number < -512.0 || number > 511.0)
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

        string binaryString;

        // Handle negative numbers using two's complement
        if (intValue < 0)
        {
            // Convert to 10-bit two's complement
            int twosComplement = 1024 + intValue; // 2^10
            binaryString = Convert.ToString(twosComplement, 2);
        }
        else
        {
            binaryString = Convert.ToString(intValue, 2);
        }

        // Apply padding if places specified
        if (places > 0)
        {
            if (binaryString.Length > places)
            {
                return CellValue.Error("#NUM!");
            }

            binaryString = binaryString.PadLeft(places, '0');
        }

        return CellValue.FromString(binaryString);
    }
}
