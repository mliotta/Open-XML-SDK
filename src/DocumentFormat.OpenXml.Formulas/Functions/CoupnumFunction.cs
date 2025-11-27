// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COUPNUM function.
/// COUPNUM(settlement, maturity, frequency, [basis]) - returns the number of coupons payable between settlement and maturity.
/// </summary>
public sealed class CoupnumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CoupnumFunction Instance = new();

    private CoupnumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COUPNUM";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // Check for errors in all provided arguments first (Excel error propagation rule)
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        if (args.Length < 3 || args.Length > 4)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check types for required arguments
        if (args[0].Type != FormulaResultType.Number ||
            args[1].Type != FormulaResultType.Number ||
            args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var basis = 0;
        if (args.Length == 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == FormulaResultType.Number)
            {
                basis = (int)args[3].NumericValue;
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
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var frequency = (int)args[2].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return FormulaResult.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return FormulaResult.Error("#NUM!");
            }

            var couponCount = DayCountHelper.CountCoupons(settlement, maturity, frequency);

            return FormulaResult.FromNumber(couponCount);
        }
        catch
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
