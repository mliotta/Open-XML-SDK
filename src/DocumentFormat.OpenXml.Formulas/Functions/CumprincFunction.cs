// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CUMPRINC function.
/// CUMPRINC(rate, nper, pv, start_period, end_period, type) - calculates the cumulative principal paid between two periods.
/// </summary>
public sealed class CumprincFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CumprincFunction Instance = new();

    private CumprincFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CUMPRINC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var rate = args[0].NumericValue;
        var nper = args[1].NumericValue;
        var pv = args[2].NumericValue;
        var startPeriod = args[3].NumericValue;
        var endPeriod = args[4].NumericValue;
        var type = args[5].NumericValue;

        // Validate inputs
        if (rate <= 0 || nper <= 0 || pv <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        if (startPeriod < 1 || endPeriod < 1 || startPeriod > endPeriod)
        {
            return CellValue.Error("#NUM!");
        }

        if (type != 0.0 && type != 1.0)
        {
            return CellValue.Error("#NUM!");
        }

        if (startPeriod < 1 || endPeriod > nper)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate the payment amount using PMT formula
        double pmt;
        var pvif = System.Math.Pow(1 + rate, nper);
        pmt = -(rate * (0 + pvif * pv)) / ((1 + rate * type) * (pvif - 1));

        if (double.IsNaN(pmt) || double.IsInfinity(pmt))
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate cumulative principal by summing PPMT for each period
        double cumulativePrincipal = 0.0;

        for (int period = (int)System.Math.Ceiling(startPeriod); period <= (int)System.Math.Floor(endPeriod); period++)
        {
            double ipmt;

            if (period == 1 && type == 1.0)
            {
                // For beginning of period payments, interest in period 1 is 0
                ipmt = 0.0;
            }
            else
            {
                // Calculate remaining balance at the start of the period
                double remainingBalance;
                var periodsElapsed = type == 1.0 ? period - 2 : period - 1;

                if (periodsElapsed <= 0)
                {
                    remainingBalance = pv;
                }
                else
                {
                    var pvifElapsed = System.Math.Pow(1 + rate, periodsElapsed);
                    remainingBalance = pv * pvifElapsed + pmt * (1 + rate * type) * (pvifElapsed - 1) / rate;
                }

                // Interest for the period is the remaining balance times the rate
                ipmt = remainingBalance * rate;

                // For beginning of period, adjust
                if (type == 1.0)
                {
                    ipmt /= (1 + rate);
                }
            }

            // Principal = Payment - Interest
            double ppmt = pmt - ipmt;
            cumulativePrincipal += ppmt;
        }

        if (double.IsNaN(cumulativePrincipal) || double.IsInfinity(cumulativePrincipal))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(cumulativePrincipal);
    }
}
