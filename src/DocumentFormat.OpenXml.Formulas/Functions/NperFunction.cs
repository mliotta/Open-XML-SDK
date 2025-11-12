// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NPER function.
/// NPER(rate, pmt, pv, [fv], [type]) - calculates the number of periods for an investment based on periodic, constant payments and a constant interest rate.
/// </summary>
public sealed class NperFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NperFunction Instance = new();

    private NperFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NPER";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 5)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[2].IsError)
        {
            return args[2];
        }

        // Validate required arguments are numbers
        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rate = args[0].NumericValue;
        var pmt = args[1].NumericValue;
        var pv = args[2].NumericValue;
        var fv = 0.0;
        var type = 0.0;

        // Optional fv parameter
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            fv = args[3].NumericValue;
        }

        // Optional type parameter
        if (args.Length == 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            type = args[4].NumericValue;
        }

        // Validate type is 0 or 1
        if (type != 0.0 && type != 1.0)
        {
            return CellValue.Error("#NUM!");
        }

        // Validate pmt is not zero when rate is zero
        if (rate == 0.0 && pmt == 0.0)
        {
            return CellValue.Error("#NUM!");
        }

        double nper;

        // Special case: rate = 0
        if (rate == 0.0)
        {
            nper = -(pv + fv) / pmt;
        }
        else
        {
            // Standard NPER formula
            var pmtWithType = pmt * (1 + rate * type);

            // Check for valid inputs to avoid log of negative number
            if (pmtWithType == 0.0)
            {
                return CellValue.Error("#NUM!");
            }

            var numerator = pmtWithType - fv * rate;
            var denominator = pmtWithType + pv * rate;

            if (numerator <= 0.0 || denominator <= 0.0)
            {
                return CellValue.Error("#NUM!");
            }

            nper = System.Math.Log(numerator / denominator) / System.Math.Log(1 + rate);
        }

        if (double.IsNaN(nper) || double.IsInfinity(nper) || nper < 0.0)
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(nper);
    }
}
