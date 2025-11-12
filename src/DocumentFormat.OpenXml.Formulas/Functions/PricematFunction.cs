// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PRICEMAT function.
/// PRICEMAT(settlement, maturity, issue, rate, yld, [basis]) - returns the price per $100 face value of a security that pays interest at maturity.
/// </summary>
public sealed class PricematFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PricematFunction Instance = new();

    private PricematFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PRICEMAT";

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
            var issue = DateTime.FromOADate(args[2].NumericValue);
            var rate = args[3].NumericValue;
            var yld = args[4].NumericValue;

            // Validate inputs
            if (settlement >= maturity || issue >= settlement || rate < 0 || yld < 0)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate year fractions
            var issueToSettlement = DayCountHelper.DayCountFraction(issue, settlement, basis);
            var issueToMaturity = DayCountHelper.DayCountFraction(issue, maturity, basis);
            var settlementToMaturity = DayCountHelper.DayCountFraction(settlement, maturity, basis);

            // PRICEMAT formula
            var numerator = 100 + (100 * rate * issueToMaturity);
            var denominator = 1 + (yld * settlementToMaturity);
            var accruedInterest = 100 * rate * issueToSettlement;

            var price = (numerator / denominator) - accruedInterest;

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
