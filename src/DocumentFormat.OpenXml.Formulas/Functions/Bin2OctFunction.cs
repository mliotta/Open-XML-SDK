// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BIN2OCT function.
/// BIN2OCT(number, [places]) - converts binary to octal.
/// </summary>
public sealed class Bin2OctFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Bin2OctFunction Instance = new();

    private Bin2OctFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BIN2OCT";

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

        var binaryString = args[0].StringValue.Trim();

        // Validate binary string length (max 10 characters for 10-bit)
        if (binaryString.Length > 10)
        {
            return CellValue.Error("#NUM!");
        }

        // Validate binary string contains only 0s and 1s
        foreach (char c in binaryString)
        {
            if (c != '0' && c != '1')
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
            int decimalValue;

            // Handle negative numbers (two's complement for 10-bit)
            if (binaryString.Length == 10 && binaryString[0] == '1')
            {
                // Negative number in two's complement
                decimalValue = Convert.ToInt32(binaryString, 2);
                decimalValue = decimalValue - 1024; // 2^10
            }
            else
            {
                decimalValue = Convert.ToInt32(binaryString, 2);
            }

            string octalString;

            // Handle negative numbers using two's complement for octal (30-bit)
            if (decimalValue < 0)
            {
                // Convert to 30-bit two's complement
                long twosComplement = 0x40000000L + decimalValue;
                octalString = Convert.ToString(twosComplement, 8).ToUpperInvariant();
            }
            else
            {
                octalString = Convert.ToString(decimalValue, 8).ToUpperInvariant();
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
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
