// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PRICE function.
/// PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis]) - returns the price per $100 face value of a security that pays periodic interest.
/// </summary>
public sealed class PriceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PriceFunction Instance = new();

    private PriceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PRICE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 6 || args.Length > 7)
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
        if (args.Length == 7)
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

        try
        {
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var rate = args[2].NumericValue;
            var yld = args[3].NumericValue;
            var redemption = args[4].NumericValue;
            var frequency = (int)args[5].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity || rate < 0 || yld < 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            var couponRate = rate / frequency;
            var yieldRate = yld / frequency;
            var numCoupons = DayCountHelper.CountCoupons(settlement, maturity, frequency);

            // Days from settlement to next coupon
            var previousCoupon = DayCountHelper.GetPreviousCouponDate(settlement, maturity, frequency);
            var nextCoupon = DayCountHelper.GetNextCouponDate(settlement, maturity, frequency);
            var daysInCouponPeriod = DayCountHelper.DaysBetween(previousCoupon, nextCoupon, basis);
            var daysFromSettlement = DayCountHelper.DaysBetween(settlement, nextCoupon, basis);

            // A = Accrued days (DSC = days from settlement to next coupon)
            var dsc = daysFromSettlement;
            var e = daysInCouponPeriod;
            var a = e - dsc;

            // Calculate price using standard bond pricing formula
            double presentValue = 0.0;

            // Present value of coupons
            for (int k = 1; k <= numCoupons; k++)
            {
                var exponent = k - 1 + (dsc / e);
                var discount = System.Math.Pow(1 + yieldRate, exponent);
                presentValue += (100 * couponRate) / discount;
            }

            // Present value of redemption
            var redemptionExponent = numCoupons - 1 + (dsc / e);
            var redemptionDiscount = System.Math.Pow(1 + yieldRate, redemptionExponent);
            presentValue += redemption / redemptionDiscount;

            // Subtract accrued interest
            var accruedInterest = (100 * couponRate) * (a / e);
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
