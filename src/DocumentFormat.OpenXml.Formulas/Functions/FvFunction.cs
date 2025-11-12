// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FV function.
/// FV(rate, nper, pmt, [pv], [type]) - calculates the future value of an investment based on periodic, constant payments and a constant interest rate.
/// </summary>
public sealed class FvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FvFunction Instance = new();

    private FvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FV";

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
        var nper = args[1].NumericValue;
        var pmt = args[2].NumericValue;
        var pv = 0.0;
        var type = 0.0;

        // Optional pv parameter
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

            pv = args[3].NumericValue;
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

        // Validate nper is positive
        if (nper <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        double fv;

        // Special case: rate = 0
        if (rate == 0.0)
        {
            fv = -(pv + pmt * nper);
        }
        else
        {
            // Standard FV formula
            var pvif = System.Math.Pow(1 + rate, nper);
            fv = -(pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate);
        }

        if (double.IsNaN(fv) || double.IsInfinity(fv))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(fv);
    }
}
