// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RECEIVED function.
/// RECEIVED(settlement, maturity, investment, discount, [basis]) - returns the amount received at maturity for a fully invested security.
/// </summary>
public sealed class ReceivedFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ReceivedFunction Instance = new();

    private ReceivedFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RECEIVED";

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
            var discount = args[3].NumericValue;

            // Validate inputs
            if (investment <= 0 || discount <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            if (settlement >= maturity)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate amount received
            var dayCount = DayCountHelper.DayCountFraction(settlement, maturity, basis);
            var received = investment / (1 - (discount * dayCount));

            if (double.IsNaN(received) || double.IsInfinity(received))
            {
                return CellValue.Error("#NUM!");
            }

            return CellValue.FromNumber(received);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }
    }
}
