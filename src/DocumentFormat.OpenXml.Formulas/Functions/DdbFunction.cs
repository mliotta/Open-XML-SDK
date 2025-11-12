// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DDB function.
/// DDB(cost, salvage, life, period, [factor]) - calculates depreciation using double-declining balance method.
/// </summary>
public sealed class DdbFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DdbFunction Instance = new();

    private DdbFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DDB";

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
        var factor = 2.0;

        // Optional factor parameter
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

            factor = args[4].NumericValue;
        }

        // Validate inputs
        if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || period > life || factor <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Special case: if salvage >= cost, no depreciation
        if (salvage >= cost)
        {
            return CellValue.FromNumber(0.0);
        }

        // Calculate depreciation rate
        var rate = factor / life;

        double bookValue = cost;
        double periodDepreciation = 0.0;

        // Calculate cumulative depreciation up to the period
        for (int i = 1; i <= period; i++)
        {
            // Calculate declining balance depreciation for this period
            double ddbDepreciation = bookValue * rate;

            // Calculate straight-line depreciation for remaining periods
            double remainingLife = life - i + 1;
            double slnDepreciation = (bookValue - salvage) / remainingLife;

            // Use the larger of declining balance or straight-line
            double currentDepreciation = System.Math.Max(ddbDepreciation, slnDepreciation);

            // Ensure we don't depreciate below salvage value
            if (bookValue - currentDepreciation < salvage)
            {
                currentDepreciation = bookValue - salvage;
            }

            if (i == period)
            {
                periodDepreciation = currentDepreciation;
            }

            bookValue -= currentDepreciation;
        }

        if (double.IsNaN(periodDepreciation) || double.IsInfinity(periodDepreciation))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(periodDepreciation);
    }
}
