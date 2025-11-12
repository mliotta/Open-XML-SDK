// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DISC function.
/// DISC(settlement, maturity, pr, redemption, [basis]) - returns the discount rate for a security.
/// </summary>
public sealed class DiscFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DiscFunction Instance = new();

    private DiscFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DISC";

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

        var basis = 0;
        if (args.Length == 5 && args[4].Type != CellValueType.Empty)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == CellValueType.Number)
            {
                basis = (int)args[4].NumericValue;
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
            var pr = args[2].NumericValue;
            var redemption = args[3].NumericValue;

            // Validate inputs
            if (pr <= 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate discount rate
            var dayCount = DayCountHelper.DayCountFraction(settlement, maturity, basis);
            var discountRate = ((redemption - pr) / redemption) / dayCount;

            if (double.IsNaN(discountRate) || double.IsInfinity(discountRate))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(discountRate);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
