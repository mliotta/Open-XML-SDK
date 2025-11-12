// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AMORDEGRC function.
/// AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis]) - returns the depreciation for each accounting period using declining balance with coefficients (French accounting).
/// </summary>
public sealed class AmordegrcFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AmordegrcFunction Instance = new();

    private AmordegrcFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AMORDEGRC";

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
        if (args.Length == 7 && args[6].Type != CellValueType.Empty)
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
            var cost = args[0].NumericValue;
            var datePurchased = DateTime.FromOADate(args[1].NumericValue);
            var firstPeriod = DateTime.FromOADate(args[2].NumericValue);
            var salvage = args[3].NumericValue;
            var period = (int)args[4].NumericValue;
            var rate = args[5].NumericValue;

            // Validate inputs
            if (cost < 0 || salvage < 0 || rate <= 0 || period < 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (salvage >= cost)
            {
                return CellValue.FromNumber(0);
            }

            if (datePurchased > firstPeriod)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate total life in years
            var life = 1.0 / rate;

            // Determine the depreciation coefficient based on life
            double coefficient;
            if (life >= 0 && life < 3)
            {
                coefficient = 1.0;
            }
            else if (life >= 3 && life < 5)
            {
                coefficient = 1.5;
            }
            else if (life >= 5 && life < 6)
            {
                coefficient = 2.0;
            }
            else
            {
                coefficient = 2.5;
            }

            var depreciationRate = rate * coefficient;
            var bookValue = cost;

            // Calculate depreciation for period 0 (pro-rated first period)
            if (period == 0)
            {
                var fraction = DayCountHelper.DayCountFraction(datePurchased, firstPeriod, basis);
                var depreciation = bookValue * depreciationRate * fraction;
                var maxDepreciation = bookValue - salvage;
                return CellValue.FromNumber(System.Math.Min(depreciation, maxDepreciation));
            }

            // Calculate book value at start of requested period
            // Period 0
            var period0Fraction = DayCountHelper.DayCountFraction(datePurchased, firstPeriod, basis);
            var period0Depreciation = bookValue * depreciationRate * period0Fraction;
            bookValue -= System.Math.Min(period0Depreciation, bookValue - salvage);

            // Periods 1 to period-1
            for (int p = 1; p < period; p++)
            {
                if (bookValue <= salvage)
                {
                    break;
                }

                // Check if we should switch to linear depreciation
                var remainingLife = life - p;
                var decliningBalanceDepreciation = bookValue * depreciationRate;
                var linearDepreciation = remainingLife > 0 ? (bookValue - salvage) / remainingLife : 0;

                var periodDepreciation = System.Math.Max(decliningBalanceDepreciation, linearDepreciation);
                periodDepreciation = System.Math.Min(periodDepreciation, bookValue - salvage);

                bookValue -= periodDepreciation;
            }

            // Calculate depreciation for requested period
            if (bookValue <= salvage)
            {
                return CellValue.FromNumber(0);
            }

            var remainingLifeAtPeriod = life - period;
            var dbDepreciation = bookValue * depreciationRate;
            var slDepreciation = remainingLifeAtPeriod > 0 ? (bookValue - salvage) / remainingLifeAtPeriod : 0;

            var finalDepreciation = System.Math.Max(dbDepreciation, slDepreciation);
            finalDepreciation = System.Math.Min(finalDepreciation, bookValue - salvage);

            if (double.IsNaN(finalDepreciation) || double.IsInfinity(finalDepreciation))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(System.Math.Max(0, finalDepreciation));
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
