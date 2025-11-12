// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the AMORLINC function.
/// AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis]) - returns the depreciation for each accounting period using linear depreciation (French accounting).
/// </summary>
public sealed class AmorlincFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AmorlincFunction Instance = new();

    private AmorlincFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "AMORLINC";

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
            var depreciableAmount = cost - salvage;
            var annualDepreciation = depreciableAmount * rate;

            // For period 0, calculate pro-rated depreciation from purchase to first period end
            if (period == 0)
            {
                var fraction = DayCountHelper.DayCountFraction(datePurchased, firstPeriod, basis);
                var depreciation = cost * rate * fraction;
                return CellValue.FromNumber(System.Math.Min(depreciation, depreciableAmount));
            }

            // Calculate accumulated depreciation before this period
            var accumulatedDepreciation = 0.0;

            // Period 0 depreciation
            var period0Fraction = DayCountHelper.DayCountFraction(datePurchased, firstPeriod, basis);
            accumulatedDepreciation = cost * rate * period0Fraction;

            // Periods 1 to period-1
            for (int p = 1; p < period; p++)
            {
                accumulatedDepreciation += annualDepreciation;
            }

            // Check if already fully depreciated
            if (accumulatedDepreciation >= depreciableAmount)
            {
                return CellValue.FromNumber(0);
            }

            // Calculate depreciation for the requested period
            var periodDepreciation = System.Math.Min(annualDepreciation, depreciableAmount - accumulatedDepreciation);

            if (double.IsNaN(periodDepreciation) || double.IsInfinity(periodDepreciation))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(System.Math.Max(0, periodDepreciation));
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
