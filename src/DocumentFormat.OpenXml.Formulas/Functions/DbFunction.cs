// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DB function.
/// DB(cost, salvage, life, period, [month]) - calculates depreciation using fixed-declining balance method.
/// </summary>
public sealed class DbFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DbFunction Instance = new();

    private DbFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 4 || args.Length > 5)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < 4; i++)
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

        var cost = args[0].NumericValue;
        var salvage = args[1].NumericValue;
        var life = args[2].NumericValue;
        var period = args[3].NumericValue;
        var month = 12.0;

        // Optional month parameter
        if (args.Length == 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            month = args[4].NumericValue;
        }

        // Validate inputs
        if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || period > life || month < 1 || month > 12)
        {
            return CellValue.Error("#NUM!");
        }

        // Special case: if salvage >= cost, no depreciation
        if (salvage >= cost)
        {
            return CellValue.FromNumber(0.0);
        }

        // Calculate fixed-declining balance rate: rate = 1 - ((salvage/cost)^(1/life))
        var rate = 1.0 - System.Math.Pow(salvage / cost, 1.0 / life);

        // Round rate to 3 decimal places (Excel behavior)
        rate = System.Math.Round(rate, 3);

        double totalDepreciation = 0.0;
        double bookValue = cost;

        // Calculate depreciation for the specified period
        for (int i = 1; i <= period; i++)
        {
            double periodDepreciation;

            if (i == 1)
            {
                // First period uses the month parameter
                periodDepreciation = cost * rate * month / 12.0;
            }
            else if (i == life + 1)
            {
                // Last period (if partial year in first period)
                periodDepreciation = bookValue * rate * (12.0 - month) / 12.0;
            }
            else
            {
                // Full year depreciation
                periodDepreciation = bookValue * rate;
            }

            if (i == period)
            {
                totalDepreciation = periodDepreciation;
            }

            bookValue -= periodDepreciation;
        }

        if (double.IsNaN(totalDepreciation) || double.IsInfinity(totalDepreciation))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(totalDepreciation);
    }
}
