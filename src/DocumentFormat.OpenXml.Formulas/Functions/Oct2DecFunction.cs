// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the OCT2DEC function.
/// OCT2DEC(number) - converts octal to decimal.
/// </summary>
public sealed class Oct2DecFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Oct2DecFunction Instance = new();

    private Oct2DecFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "OCT2DEC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
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

        // Validate octal string contains only valid characters (0-7)
        foreach (char c in octalString)
        {
            if (c < '0' || c > '7')
            {
                return CellValue.Error("#NUM!");
            }
        }

        try
        {
            // Handle negative numbers (two's complement for 30-bit)
            if (octalString.Length == 10)
            {
                long value = Convert.ToInt64(octalString, 8);
                // Check if this represents a negative number (bit 29 set)
                if (value >= 0x20000000L)
                {
                    value = value - 0x40000000L;
                }

                return CellValue.FromNumber(value);
            }
            else
            {
                int value = Convert.ToInt32(octalString, 8);
                return CellValue.FromNumber(value);
            }
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
