// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ODDFPRICE function.
/// ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, frequency, [basis]) - returns the price per $100 face value of a security with an odd first period.
/// </summary>
public sealed class OddfpriceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OddfpriceFunction Instance = new();

    private OddfpriceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ODDFPRICE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 8 || args.Length > 9)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 8; i++)
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
        if (args.Length == 9 && args[8].Type != CellValueType.Empty)
        {
            if (args[8].IsError)
            {
                return args[8];
            }

            if (args[8].Type == CellValueType.Number)
            {
                basis = (int)args[8].NumericValue;
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
            var issue = DateTime.FromOADate(args[2].NumericValue);
            var firstCoupon = DateTime.FromOADate(args[3].NumericValue);
            var rate = args[4].NumericValue;
            var yld = args[5].NumericValue;
            var redemption = args[6].NumericValue;
            var frequency = (int)args[7].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (rate < 0 || yld < 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (issue >= settlement || settlement >= firstCoupon || firstCoupon >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate number of coupons from first coupon to maturity (excluding the odd first period)
            var numCoupons = DayCountHelper.CountCoupons(firstCoupon, maturity, frequency);

            // Days and calculations for the odd first period
            var dsc = DayCountHelper.DaysBetween(settlement, firstCoupon, basis);
            var e = DayCountHelper.DaysBetween(issue, firstCoupon, basis);
            var a = e - dsc;

            var couponRate = rate / frequency;
            var yieldRate = yld / frequency;

            // Calculate present value
            double presentValue = 0.0;

            // First coupon payment (odd period)
            var firstCouponPayment = 100 * couponRate * (e / DayCountHelper.DaysBetween(
                DayCountHelper.GetPreviousCouponDate(firstCoupon, maturity, frequency),
                firstCoupon,
                basis));
            var firstDiscount = System.Math.Pow(1 + yieldRate, dsc / DayCountHelper.DaysBetween(settlement, firstCoupon, basis));
            presentValue += firstCouponPayment / firstDiscount;

            // Regular coupon payments
            for (int k = 2; k <= numCoupons; k++)
            {
                var discount = System.Math.Pow(1 + yieldRate, k - 1 + dsc / e);
                presentValue += (100 * couponRate) / discount;
            }

            // Redemption value
            var redemptionDiscount = System.Math.Pow(1 + yieldRate, numCoupons - 1 + dsc / e);
            presentValue += redemption / redemptionDiscount;

            // Subtract accrued interest
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
