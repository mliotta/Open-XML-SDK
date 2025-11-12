// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NPV function.
/// NPV(rate, value1, [value2], ...) - calculates the net present value of an investment based on a discount rate and a series of future cash flows.
/// </summary>
public sealed class NpvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NpvFunction Instance = new();

    private NpvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NPV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
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

        // NPV formula: Σ(value[i] / (1 + rate)^i) for i=1 to n
        double npv = 0.0;

        for (int i = 1; i < args.Length; i++)
        {
            // Check for errors in value arguments
            if (args[i].IsError)
            {
                return args[i];
            }

            // Validate value argument is a number
            if (args[i].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            var value = args[i].NumericValue;
            var discountFactor = System.Math.Pow(1 + rate, i);

            if (double.IsInfinity(discountFactor) || double.IsNaN(discountFactor))
            {
                return CellValue.Error("#NUM!");
            }

            npv += value / discountFactor;
        }

        if (double.IsNaN(npv) || double.IsInfinity(npv))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(npv);
    }
}
