// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ACCRINT function.
/// ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]) - returns the accrued interest for a security that pays periodic interest.
/// </summary>
public sealed class AccrintFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AccrintFunction Instance = new();

    private AccrintFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ACCRINT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 6 || args.Length > 8)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 6; i++)
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

        var basis = 0;
        if (args.Length >= 7 && args[6].Type != CellValueType.Empty)
        {
            if (args[6].IsError)
            {
                return args[6];
            }

            if (args[6].Type == CellValueType.Number)
            {
                basis = (int)args[6].NumericValue;
                if (!DayCountHelper.IsValidBasis(basis))
                {
                    return CellValue.Error("#NUM!");
                }
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var calcMethod = 1;
        if (args.Length == 8 && args[7].Type != CellValueType.Empty)
        {
            if (args[7].IsError)
            {
                return args[7];
            }

            if (args[7].Type == CellValueType.Number)
            {
                calcMethod = (int)args[7].NumericValue;
                if (calcMethod != 0 && calcMethod != 1)
                {
                    return CellValue.Error("#NUM!");
                }
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        try
        {
            var issue = DateTime.FromOADate(args[0].NumericValue);
            var firstInterest = DateTime.FromOADate(args[1].NumericValue);
            var settlement = DateTime.FromOADate(args[2].NumericValue);
            var rate = args[3].NumericValue;
            var par = args[4].NumericValue;
            var frequency = (int)args[5].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (rate <= 0 || par <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (issue >= settlement || settlement >= firstInterest)
            {
                return CellValue.Error("#NUM!");
            }

            double accruedInterest;

            if (calcMethod == 1)
            {
                // Method 1: Calculate from issue to settlement
                var dayCount = DayCountHelper.DayCountFraction(issue, settlement, basis);
                accruedInterest = par * rate * dayCount;
            }
            else
            {
                // Method 0: Calculate quasi-coupon periods
                var monthsPerPeriod = 12 / frequency;
                var numCoupons = 0;
                var currentDate = firstInterest;

                // Work backwards from first interest to find coupon dates before settlement
                while (currentDate > settlement)
                {
                    currentDate = currentDate.AddMonths(-monthsPerPeriod);
                    numCoupons++;
                }

                // Calculate accrued interest by summing up coupon periods
                accruedInterest = 0.0;
                var periodStart = issue;

                for (int i = 0; i < numCoupons; i++)
                {
                    var periodEnd = periodStart.AddMonths(monthsPerPeriod);
                    if (periodEnd > settlement)
                    {
                        periodEnd = settlement;
                    }

                    var dayCount = DayCountHelper.DayCountFraction(periodStart, periodEnd, basis);
                    accruedInterest += par * rate * dayCount;

                    periodStart = periodStart.AddMonths(monthsPerPeriod);
                    if (periodStart >= settlement)
                    {
                        break;
                    }
                }
            }

            if (double.IsNaN(accruedInterest) || double.IsInfinity(accruedInterest))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(accruedInterest);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
