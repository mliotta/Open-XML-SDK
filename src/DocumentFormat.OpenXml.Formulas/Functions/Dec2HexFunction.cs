// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DEC2HEX function.
/// DEC2HEX(number, [places]) - converts decimal to hexadecimal.
/// </summary>
public sealed class Dec2HexFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly Dec2HexFunction Instance = new();

    private Dec2HexFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DEC2HEX";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // Validate range: -549755813888 to 549755813887 (40-bit signed)
        if (number < -549755813888.0 || number > 549755813887.0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Truncate to integer
        var intValue = (long)System.Math.Floor(number);

        int places = 0;
        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            places = (int)System.Math.Floor(args[1].NumericValue);

            if (places < 0)
            {
                return FormulaResult.Error("#NUM!");
            }

            if (places > 10)
            {
                return FormulaResult.Error("#NUM!");
            }
        }

        string hexString;

        // Handle negative numbers using two's complement
        if (intValue < 0)
        {
            // Convert to 40-bit two's complement
            long twosComplement = 0x10000000000L + intValue;
            hexString = twosComplement.ToString("X", CultureInfo.InvariantCulture);
        }
        else
        {
            hexString = intValue.ToString("X", CultureInfo.InvariantCulture);
        }

        // Apply padding if places specified
        if (places > 0)
        {
            if (hexString.Length > places)
            {
                return FormulaResult.Error("#NUM!");
            }

            hexString = hexString.PadLeft(places, '0');
        }

        return FormulaResult.FromString(hexString);
    }
}
