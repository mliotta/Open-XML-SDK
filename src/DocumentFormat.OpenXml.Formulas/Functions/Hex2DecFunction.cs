// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HEX2DEC function.
/// HEX2DEC(number) - converts hexadecimal to decimal.
/// </summary>
public sealed class Hex2DecFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Hex2DecFunction Instance = new();

    private Hex2DecFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HEX2DEC";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        var hexString = args[0].StringValue.Trim();

        // Validate hex string length (max 10 characters for signed 40-bit)
        if (hexString.Length > 10)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Validate hex string contains only valid characters
        foreach (char c in hexString)
        {
            if (!IsHexChar(c))
            {
                return FormulaResult.Error("#NUM!");
            }
        }

        try
        {
            // Handle negative numbers (two's complement)
            if (hexString.Length == 10)
            {
                long value = Convert.ToInt64(hexString, 16);
                // Check if this represents a negative number (bit 39 set)
                if (value >= 0x8000000000L)
                {
                    value = value - 0x10000000000L;
                }

                return FormulaResult.FromNumber(value);
            }
            else
            {
                int value = Convert.ToInt32(hexString, 16);
                return FormulaResult.FromNumber(value);
            }
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }

    private static bool IsHexChar(char c)
    {
        return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
    }
}
