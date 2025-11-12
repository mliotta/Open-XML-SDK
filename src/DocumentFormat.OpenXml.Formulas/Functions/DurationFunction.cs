// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DURATION function.
/// DURATION(settlement, maturity, coupon, yld, frequency, [basis]) - returns the Macaulay duration for an assumed par value of $100.
/// </summary>
public sealed class DurationFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DurationFunction Instance = new();

    private DurationFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DURATION";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 5 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 5; i++)
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
        if (args.Length == 6)
        {
            if (args[5].IsError)
            {
                return args[5];
            }

            if (args[5].Type == CellValueType.Number)
            {
                basis = (int)args[5].NumericValue;
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
            var coupon = args[2].NumericValue;
            var yld = args[3].NumericValue;
            var frequency = (int)args[4].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity || coupon < 0 || yld < 0)
            {
                return CellValue.Error("#NUM!");
            }

            var couponRate = coupon / frequency;
            var yieldRate = yld / frequency;
            var numCoupons = DayCountHelper.CountCoupons(settlement, maturity, frequency);

            var previousCoupon = DayCountHelper.GetPreviousCouponDate(settlement, maturity, frequency);
            var nextCoupon = DayCountHelper.GetNextCouponDate(settlement, maturity, frequency);
            var daysInCouponPeriod = DayCountHelper.DaysBetween(previousCoupon, nextCoupon, basis);
            var daysFromSettlement = DayCountHelper.DaysBetween(settlement, nextCoupon, basis);

            var dsc = daysFromSettlement;
            var e = daysInCouponPeriod;

            // Calculate weighted present value of cash flows
            double weightedPV = 0.0;
            double totalPV = 0.0;

            // Present value of coupons
            for (int k = 1; k <= numCoupons; k++)
            {
                var timeToCoupon = (k - 1 + (dsc / e)) / frequency;
                var exponent = k - 1 + (dsc / e);
                var discount = System.Math.Pow(1 + yieldRate, exponent);
                var cashFlow = 100 * couponRate;
                var pv = cashFlow / discount;

                weightedPV += timeToCoupon * pv;
                totalPV += pv;
            }

            // Present value of redemption
            var timeToRedemption = (numCoupons - 1 + (dsc / e)) / frequency;
            var redemptionExponent = numCoupons - 1 + (dsc / e);
            var redemptionDiscount = System.Math.Pow(1 + yieldRate, redemptionExponent);
            var redemptionPV = 100.0 / redemptionDiscount;

            weightedPV += timeToRedemption * redemptionPV;
            totalPV += redemptionPV;

            // Macaulay duration
            var duration = weightedPV / totalPV;

            if (double.IsNaN(duration) || double.IsInfinity(duration) || duration < 0)
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(duration);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
