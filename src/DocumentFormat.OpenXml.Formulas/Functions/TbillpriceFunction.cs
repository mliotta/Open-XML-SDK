// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TBILLPRICE function.
/// TBILLPRICE(settlement, maturity, discount) - returns the price per $100 face value for a Treasury bill.
/// </summary>
public sealed class TbillpriceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TbillpriceFunction Instance = new();

    private TbillpriceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TBILLPRICE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        for (int i = 0; i < 3; i++)
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

        try
        {
            var settlement = DateTime.FromOADate(args[0].NumericValue);
            var maturity = DateTime.FromOADate(args[1].NumericValue);
            var discount = args[2].NumericValue;

            // Validate inputs
            if (discount <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // T-bills must mature within one year
            var daysToMaturity = (maturity - settlement).TotalDays;
            if (daysToMaturity > 366)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate price using actual/360 convention
            var price = 100 * (1 - discount * daysToMaturity / 360.0);

            if (double.IsNaN(price) || double.IsInfinity(price) || price <= 0)
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
