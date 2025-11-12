// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VDB function.
/// VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch]) - returns the depreciation of an asset for any period using the variable declining balance method.
/// </summary>
public sealed class VdbFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly VdbFunction Instance = new();

    private VdbFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VDB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 5 || args.Length > 7)
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

        var factor = 2.0;
        if (args.Length >= 6 && args[5].Type != CellValueType.Empty)
        {
            if (args[5].IsError)
            {
                return args[5];
            }

            if (args[5].Type == CellValueType.Number)
            {
                factor = args[5].NumericValue;
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var noSwitch = false;
        if (args.Length == 7 && args[6].Type != CellValueType.Empty)
        {
            if (args[6].IsError)
            {
                return args[6];
            }

            if (args[6].Type == CellValueType.Boolean)
            {
                noSwitch = args[6].BoolValue;
            }
            else if (args[6].Type == CellValueType.Number)
            {
                noSwitch = args[6].NumericValue != 0;
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        try
        {
            var cost = args[0].NumericValue;
            var salvage = args[1].NumericValue;
            var life = args[2].NumericValue;
            var startPeriod = args[3].NumericValue;
            var endPeriod = args[4].NumericValue;

            // Validate inputs
            if (cost < 0 || salvage < 0 || life <= 0 || factor <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (startPeriod < 0 || endPeriod < startPeriod || endPeriod > life)
            {
                return CellValue.Error("#NUM!");
            }

            if (salvage >= cost)
            {
                return CellValue.FromNumber(0);
            }

            // Calculate depreciation
            var totalDepreciation = 0.0;
            var bookValue = cost;
            var ddbRate = factor / life;

            // Calculate depreciation for each period from start to end
            for (var period = System.Math.Floor(startPeriod); period <= System.Math.Ceiling(endPeriod); period++)
            {
                var periodStart = System.Math.Max(startPeriod, period);
                var periodEnd = System.Math.Min(endPeriod, period + 1);
                var periodFraction = periodEnd - periodStart;

                if (periodFraction <= 0)
                {
                    continue;
                }

                // Calculate book value at the start of this period
                bookValue = cost;
                for (var p = 0; p < period; p++)
                {
                    var depreciation = CalculatePeriodDepreciation(bookValue, salvage, life, p, ddbRate, noSwitch);
                    bookValue -= depreciation;
                    if (bookValue < salvage)
                    {
                        bookValue = salvage;
                        break;
                    }
                }

                // Calculate depreciation for this period
                var periodDepreciation = CalculatePeriodDepreciation(bookValue, salvage, life, period, ddbRate, noSwitch);
                totalDepreciation += periodDepreciation * periodFraction;
            }

            if (double.IsNaN(totalDepreciation) || double.IsInfinity(totalDepreciation))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(totalDepreciation);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }

    private static double CalculatePeriodDepreciation(double bookValue, double salvage, double life, double period, double ddbRate, bool noSwitch)
    {
        if (bookValue <= salvage)
        {
            return 0;
        }

        // Calculate DDB depreciation
        var ddbDepreciation = bookValue * ddbRate;

        // Calculate straight-line depreciation for remaining life
        var remainingLife = life - period;
        var slDepreciation = remainingLife > 0 ? (bookValue - salvage) / remainingLife : 0;

        // If no_switch is true, always use DDB
        if (noSwitch)
        {
            var depreciation = System.Math.Min(ddbDepreciation, bookValue - salvage);
            return System.Math.Max(0, depreciation);
        }

        // Otherwise, use the greater of DDB or straight-line
        var useDepreciation = System.Math.Max(ddbDepreciation, slDepreciation);
        useDepreciation = System.Math.Min(useDepreciation, bookValue - salvage);
        return System.Math.Max(0, useDepreciation);
    }
}
