// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the OCT2BIN function.
/// OCT2BIN(number, [places]) - converts octal to binary.
/// </summary>
public sealed class Oct2BinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Oct2BinFunction Instance = new();

    private Oct2BinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "OCT2BIN";

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

        var octalString = args[0].StringValue.Trim();

        // Validate octal string length (max 10 characters for 30-bit)
        if (octalString.Length > 10)
        {
            return CellValue.Error("#NUM!");
        }

        // Validate octal string contains only 0-7
        foreach (char c in octalString)
        {
            if (c < '0' || c > '7')
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
            // Convert octal to decimal
            long decimalValue = Convert.ToInt64(octalString, 8);

            // Handle negative numbers (two's complement for 30-bit)
            if (octalString.Length == 10 && octalString[0] >= '4')
            {
                // Negative number in two's complement
                decimalValue = decimalValue - 0x40000000L;
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
