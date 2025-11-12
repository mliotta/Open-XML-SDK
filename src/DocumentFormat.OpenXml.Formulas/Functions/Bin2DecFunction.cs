// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BIN2DEC function.
/// BIN2DEC(number) - converts binary to decimal.
/// </summary>
public sealed class Bin2DecFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Bin2DecFunction Instance = new();

    private Bin2DecFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BIN2DEC";

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

        try
        {
            // Handle negative numbers (two's complement for 10-bit)
            if (binaryString.Length == 10 && binaryString[0] == '1')
            {
                // Negative number in two's complement
                int value = Convert.ToInt32(binaryString, 2);
                value = value - 1024; // 2^10
                return CellValue.FromNumber(value);
            }
            else
            {
                int value = Convert.ToInt32(binaryString, 2);
                return CellValue.FromNumber(value);
            }
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
