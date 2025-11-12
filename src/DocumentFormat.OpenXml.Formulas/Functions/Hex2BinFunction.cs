// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HEX2BIN function.
/// HEX2BIN(number, [places]) - converts hexadecimal to binary.
/// </summary>
public sealed class Hex2BinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Hex2BinFunction Instance = new();

    private Hex2BinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HEX2BIN";

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

        var hexString = args[0].StringValue.Trim();

        // Validate hex string length (max 10 characters for 40-bit)
        if (hexString.Length > 10)
        {
            return CellValue.Error("#NUM!");
        }

        // Validate hex string contains only valid hex characters
        foreach (char c in hexString)
        {
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')))
            {
                return CellValue.Error("#NUM!");
            }
        }

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

        try
        {
            // Convert hex to decimal
            long decimalValue = Convert.ToInt64(hexString, 16);

            // Handle negative numbers (two's complement for 40-bit)
            if (hexString.Length == 10 && hexString[0] >= '8')
            {
                // Negative number in two's complement
                decimalValue = decimalValue - 0x10000000000L;
            }

            // Validate range for binary output (-512 to 511 for 10-bit)
            if (decimalValue < -512 || decimalValue > 511)
            {
                return CellValue.Error("#NUM!");
            }

            string binaryString;

            // Handle negative numbers using two's complement for binary (10-bit)
            if (decimalValue < 0)
            {
                // Convert to 10-bit two's complement
                long twosComplement = 1024 + decimalValue;
                binaryString = Convert.ToString(twosComplement, 2);
            }
            else
            {
                binaryString = Convert.ToString(decimalValue, 2);
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
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
