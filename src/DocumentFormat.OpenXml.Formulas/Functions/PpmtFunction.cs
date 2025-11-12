// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PPMT function.
/// PPMT(rate, per, nper, pv, [fv], [type]) - calculates the principal payment for a given period of an investment.
/// </summary>
public sealed class PpmtFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PpmtFunction Instance = new();

    private PpmtFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PPMT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 4 || args.Length > 6)
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

        if (args[3].IsError)
        {
            return args[3];
        }

        // Validate required arguments are numbers
        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number ||
            args[2].Type != CellValueType.Number || args[3].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rate = args[0].NumericValue;
        var per = args[1].NumericValue;
        var nper = args[2].NumericValue;
        var pv = args[3].NumericValue;
        var fv = 0.0;
        var type = 0.0;

        // Optional fv parameter
        if (args.Length >= 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            fv = args[4].NumericValue;
        }

        // Optional type parameter
        if (args.Length == 6)
        {
            if (args[5].IsError)
            {
                return args[5];
            }

            if (args[5].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            type = args[5].NumericValue;
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

        // Validate per is in valid range
        if (per < 1 || per > nper)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate the payment amount first using PMT formula
        double pmt;

        if (rate == 0.0)
        {
            pmt = -(pv + fv) / nper;
        }
        else
        {
            var pvif = System.Math.Pow(1 + rate, nper);
            pmt = -(rate * (fv + pvif * pv)) / ((1 + rate * type) * (pvif - 1));
        }

        if (double.IsNaN(pmt) || double.IsInfinity(pmt))
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate the interest portion using IPMT logic
        double ipmt;

        if (rate == 0.0)
        {
            // With zero interest, there is no interest payment
            ipmt = 0.0;
        }
        else
        {
            if (per == 1 && type == 1.0)
            {
                // For beginning of period payments, interest in period 1 is 0
                ipmt = 0.0;
            }
            else
            {
                // Calculate remaining balance at the start of the period
                var periodsElapsed = type == 1.0 ? per - 2 : per - 1;

                double remainingBalance;

                if (periodsElapsed <= 0)
                {
                    remainingBalance = pv;
                }
                else
                {
                    var pvif = System.Math.Pow(1 + rate, periodsElapsed);
                    remainingBalance = pv * pvif + pmt * (1 + rate * type) * (pvif - 1) / rate;
                }

                // Interest for the period is the remaining balance times the rate
                ipmt = remainingBalance * rate;

                // For beginning of period, adjust
                if (type == 1.0)
                {
                    ipmt /= (1 + rate);
                }
            }
        }

        // Principal payment = Total payment - Interest payment
        var ppmt = pmt - ipmt;

        if (double.IsNaN(ppmt) || double.IsInfinity(ppmt))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(ppmt);
    }
}
