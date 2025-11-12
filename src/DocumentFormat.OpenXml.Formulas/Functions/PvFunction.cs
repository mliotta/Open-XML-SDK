// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PV function.
/// PV(rate, nper, pmt, [fv], [type]) - calculates the present value of an investment based on a constant interest rate.
/// </summary>
public sealed class PvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PvFunction Instance = new();

    private PvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PV";

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

        // Validate nper is positive
        if (nper <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        double pv;

        // Special case: rate = 0
        if (rate == 0.0)
        {
            pv = -(fv + pmt * nper);
        }
        else
        {
            // Standard PV formula
            var pvif = System.Math.Pow(1 + rate, nper);
            pv = -(fv + pmt * (1 + rate * type) * (pvif - 1) / rate) / pvif;
        }

        if (double.IsNaN(pv) || double.IsInfinity(pv))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(pv);
    }
}
