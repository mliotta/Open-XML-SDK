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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 6 || args.Length > 7)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 6; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        var basis = 0;
        if (args.Length == 7 && args[6].Type != FormulaResultType.Empty)
        {
            if (args[6].IsError)
            {
                return args[6];
            }

            if (args[6].Type == FormulaResultType.Number)
            {
                basis = (int)args[6].NumericValue;
                if (!DayCountHelper.IsValidBasis(basis))
                {
                    return FormulaResult.Error("#NUM!");
                }
            }
            else
            {
                return FormulaResult.Error("#VALUE!");
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
                return FormulaResult.Error("#NUM!");
            }

            if (salvage >= cost)
            {
                return FormulaResult.FromNumber(0);
            }

            if (datePurchased > firstPeriod)
            {
                return FormulaResult.Error("#NUM!");
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
                return FormulaResult.FromNumber(System.Math.Min(depreciation, depreciableAmount));
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
                return FormulaResult.FromNumber(0);
            }

            // Calculate depreciation for the requested period
            var periodDepreciation = System.Math.Min(annualDepreciation, depreciableAmount - accumulatedDepreciation);

            if (double.IsNaN(periodDepreciation) || double.IsInfinity(periodDepreciation))
            {
                return FormulaResult.Error("#NUM!");
            }

            return FormulaResult.FromNumber(System.Math.Max(0, periodDepreciation));
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
