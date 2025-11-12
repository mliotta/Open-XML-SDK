// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TBILLEQ function.
/// TBILLEQ(settlement, maturity, discount) - returns the bond-equivalent yield for a Treasury bill.
/// </summary>
public sealed class TbilleqFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TbilleqFunction Instance = new();

    private TbilleqFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TBILLEQ";

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

            // Calculate bond-equivalent yield
            double bondEquivalentYield;

            if (daysToMaturity <= 182)
            {
                // For T-bills with 182 days or less to maturity
                bondEquivalentYield = (365 * discount) / (360 - discount * daysToMaturity);
            }
            else
            {
                // For T-bills with more than 182 days to maturity
                var price = 100 * (1 - discount * daysToMaturity / 360.0);
                var term1 = -daysToMaturity / 365.0;
                var term2 = System.Math.Sqrt(System.Math.Pow(daysToMaturity / 365.0, 2) - (2 * daysToMaturity / 365.0 - 1) * (1 - 100.0 / price));
                var term3 = daysToMaturity / 365.0 - 1;
                bondEquivalentYield = 2 * (term1 + term2) / term3;
            }

            if (double.IsNaN(bondEquivalentYield) || double.IsInfinity(bondEquivalentYield))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(bondEquivalentYield);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
