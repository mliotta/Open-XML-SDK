// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ODDFYIELD function.
/// ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, frequency, [basis]) - returns the yield of a security with an odd first period.
/// </summary>
public sealed class OddfyieldFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OddfyieldFunction Instance = new();

    private OddfyieldFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ODDFYIELD";

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
            var pr = args[5].NumericValue;
            var redemption = args[6].NumericValue;
            var frequency = (int)args[7].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (rate < 0 || pr <= 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (issue >= settlement || settlement >= firstCoupon || firstCoupon >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // Use Newton-Raphson method to solve for yield
            var guess = rate; // Initial guess
            var maxIterations = 100;
            var tolerance = 1e-8;

            for (int i = 0; i < maxIterations; i++)
            {
                // Calculate price at current yield guess
                var priceArgs = new[]
                {
                    CellValue.FromNumber(settlement.ToOADate()),
                    CellValue.FromNumber(maturity.ToOADate()),
                    CellValue.FromNumber(issue.ToOADate()),
                    CellValue.FromNumber(firstCoupon.ToOADate()),
                    CellValue.FromNumber(rate),
                    CellValue.FromNumber(guess),
                    CellValue.FromNumber(redemption),
                    CellValue.FromNumber(frequency),
                    CellValue.FromNumber(basis),
                };

                var priceResult = OddfpriceFunction.Instance.Execute(context, priceArgs);
                if (priceResult.IsError)
                {
                    return priceResult;
                }

                var calculatedPrice = priceResult.NumericValue;
                var priceDiff = calculatedPrice - pr;

                // Check for convergence
                if (System.Math.Abs(priceDiff) < tolerance)
                {
                    return CellValue.FromNumber(guess);
                }

                // Calculate derivative (price change for small yield change)
                var delta = 0.0001;
                var priceArgsPlus = new[]
                {
                    CellValue.FromNumber(settlement.ToOADate()),
                    CellValue.FromNumber(maturity.ToOADate()),
                    CellValue.FromNumber(issue.ToOADate()),
                    CellValue.FromNumber(firstCoupon.ToOADate()),
                    CellValue.FromNumber(rate),
                    CellValue.FromNumber(guess + delta),
                    CellValue.FromNumber(redemption),
                    CellValue.FromNumber(frequency),
                    CellValue.FromNumber(basis),
                };

                var pricePlusResult = OddfpriceFunction.Instance.Execute(context, priceArgsPlus);
                if (pricePlusResult.IsError)
                {
                    return pricePlusResult;
                }

                var derivative = (pricePlusResult.NumericValue - calculatedPrice) / delta;

                if (System.Math.Abs(derivative) < 1e-10)
                {
                    break; // Avoid division by zero
                }

                // Newton-Raphson update
                guess = guess - priceDiff / derivative;

                // Keep yield reasonable
                if (guess < -1 || guess > 10)
                {
                    return CellValue.Error("#NUM!");
                }
            }

            // If we didn't converge, return error
            return CellValue.Error("#NUM!");
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
