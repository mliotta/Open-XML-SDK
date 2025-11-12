// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ODDLPRICE function.
/// ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, [basis]) - returns the price per $100 face value of a security with an odd last period.
/// </summary>
public sealed class OddlpriceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OddlpriceFunction Instance = new();

    private OddlpriceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ODDLPRICE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 7 || args.Length > 8)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 7; i++)
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
        if (args.Length == 8 && args[7].Type != CellValueType.Empty)
        {
            if (args[7].IsError)
            {
                return args[7];
            }

            if (args[7].Type == CellValueType.Number)
            {
                basis = (int)args[7].NumericValue;
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

        try
        {
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var lastInterest = DateTime.FromOADate(args[2].NumericValue);
            var rate = args[3].NumericValue;
            var yld = args[4].NumericValue;
            var redemption = args[5].NumericValue;
            var frequency = (int)args[6].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (rate < 0 || yld < 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (lastInterest >= settlement || settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            var couponRate = rate / frequency;
            var yieldRate = yld / frequency;

            // Calculate days for the odd last period
            var dcl = DayCountHelper.DaysBetween(lastInterest, maturity, basis);
            var dsl = DayCountHelper.DaysBetween(settlement, maturity, basis);
            var a = DayCountHelper.DaysBetween(lastInterest, settlement, basis);

            // Normal coupon period length
            var monthsPerPeriod = 12 / frequency;
            var normalPeriodStart = maturity.AddMonths(-monthsPerPeriod);
            var e = DayCountHelper.DaysBetween(normalPeriodStart, maturity, basis);

            // Calculate present value components
            // Odd last coupon payment
            var oddLastCoupon = 100 * couponRate * (dcl / e);

            // Discount factor for odd last period
            var discountFactor = 1 + yieldRate * (dsl / e);

            // Present value = (Redemption + Odd Last Coupon) / Discount Factor - Accrued Interest
            var presentValue = (redemption + oddLastCoupon) / discountFactor;

            // Accrued interest for odd last period
            var accruedInterest = 100 * couponRate * (a / e);

            var price = presentValue - accruedInterest;

            if (double.IsNaN(price) || double.IsInfinity(price))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(price);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
