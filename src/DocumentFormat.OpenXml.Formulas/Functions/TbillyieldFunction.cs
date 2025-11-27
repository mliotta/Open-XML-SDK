// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TBILLYIELD function.
/// TBILLYIELD(settlement, maturity, pr) - returns the yield for a Treasury bill.
/// </summary>
public sealed class TbillyieldFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TbillyieldFunction Instance = new();

    private TbillyieldFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TBILLYIELD";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in arguments
        for (int i = 0; i < 3; i++)
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

        try
        {
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var pr = args[2].NumericValue;

            // Validate inputs
            if (pr <= 0)
            {
                return FormulaResult.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return FormulaResult.Error("#NUM!");
            }

            // T-bills must mature within one year
            var daysToMaturity = (maturity - settlement).TotalDays;
            if (daysToMaturity > 366)
            {
                return FormulaResult.Error("#NUM!");
            }

            // Calculate yield using actual/360 convention
            var yieldValue = ((100 - pr) / pr) * (360.0 / daysToMaturity);

            if (double.IsNaN(yieldValue) || double.IsInfinity(yieldValue))
            {
                return FormulaResult.Error("#NUM!");
            }

            return FormulaResult.FromNumber(yieldValue);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
