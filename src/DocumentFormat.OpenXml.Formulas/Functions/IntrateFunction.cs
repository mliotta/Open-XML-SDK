// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the INTRATE function.
/// INTRATE(settlement, maturity, investment, redemption, [basis]) - returns the interest rate for a fully invested security.
/// </summary>
public sealed class IntrateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IntrateFunction Instance = new();

    private IntrateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "INTRATE";

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
            var investment = args[2].NumericValue;
            var redemption = args[3].NumericValue;

            // Validate inputs
            if (investment <= 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate interest rate
            var dayCount = DayCountHelper.DayCountFraction(settlement, maturity, basis);
            var interestRate = ((redemption - investment) / investment) / dayCount;

            if (double.IsNaN(interestRate) || double.IsInfinity(interestRate))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(interestRate);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
