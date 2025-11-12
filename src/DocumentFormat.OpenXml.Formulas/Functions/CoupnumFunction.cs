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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        for (int i = 0; i < System.Math.Min(args.Length, 3); i++)
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
        if (args.Length == 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Number)
            {
                basis = (int)args[3].NumericValue;
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
            var frequency = (int)args[2].NumericValue;

            // Validate inputs
            if (!DayCountHelper.IsValidFrequency(frequency))
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            var couponCount = DayCountHelper.CountCoupons(settlement, maturity, frequency);

            return CellValue.FromNumber(couponCount);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
