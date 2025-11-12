// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the XNPV function.
/// XNPV(rate, values, dates) - calculates the net present value for a schedule of cash flows that is not necessarily periodic.
/// </summary>
public sealed class XnpvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly XnpvFunction Instance = new();

    private XnpvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "XNPV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // XNPV can be called with:
        // 1. Three arguments: rate, values_range, dates_range (proper Excel usage)
        // 2. Multiple arguments where rate is first, then alternating values and dates
        // For this implementation, we'll support format: rate, value1, date1, value2, date2, ...

        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in rate argument
        if (args[0].IsError)
        {
            return args[0];
        }

        // Validate rate argument is a number
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rate = args[0].NumericValue;

        // Remaining arguments should be value-date pairs
        var remainingArgs = args.Length - 1;

        // If we have exactly 2 remaining args, they might be ranges (not yet supported)
        // Otherwise, expect pairs of (value, date)
        if (remainingArgs % 2 != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var pairCount = remainingArgs / 2;
        var values = new double[pairCount];
        var dates = new double[pairCount];

        // Extract value-date pairs
        for (int i = 0; i < pairCount; i++)
        {
            var valueIdx = 1 + (i * 2);
            var dateIdx = 1 + (i * 2) + 1;

            if (args[valueIdx].IsError)
            {
                return args[valueIdx];
            }

            if (args[dateIdx].IsError)
            {
                return args[dateIdx];
            }

            if (args[valueIdx].Type != CellValueType.Number || args[dateIdx].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            values[i] = args[valueIdx].NumericValue;
            dates[i] = args[dateIdx].NumericValue;
        }

        if (pairCount == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // XNPV formula: Σ(value[i] / (1 + rate)^((date[i] - date[0]) / 365))
        var firstDate = dates[0];
        double xnpv = 0.0;

        for (int i = 0; i < pairCount; i++)
        {
            var daysDiff = dates[i] - firstDate;
            var yearFraction = daysDiff / 365.0;
            var discountFactor = System.Math.Pow(1 + rate, yearFraction);

            if (double.IsInfinity(discountFactor) || double.IsNaN(discountFactor) || discountFactor == 0)
            {
                return CellValue.Error("#NUM!");
            }

            xnpv += values[i] / discountFactor;
        }

        if (double.IsNaN(xnpv) || double.IsInfinity(xnpv))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(xnpv);
    }
}
