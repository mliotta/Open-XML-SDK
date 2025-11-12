// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PRICEDISC function.
/// PRICEDISC(settlement, maturity, discount, redemption, [basis]) - returns the price per $100 face value of a discounted security.
/// </summary>
public sealed class PricediscFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PricediscFunction Instance = new();

    private PricediscFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PRICEDISC";

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
        if (args.Length == 5)
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
            var discount = args[2].NumericValue;
            var redemption = args[3].NumericValue;

            // Validate inputs
            if (settlement >= maturity || discount <= 0 || redemption <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate fraction of year
            var yearFraction = DayCountHelper.DayCountFraction(settlement, maturity, basis);

            // PRICEDISC formula: redemption - discount * redemption * (DSM / B)
            // where DSM is days from settlement to maturity, B is year basis
            var price = redemption - (discount * redemption * yearFraction);

            if (double.IsNaN(price) || double.IsInfinity(price) || price < 0)
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(price);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
