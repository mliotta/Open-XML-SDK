// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MIRR function.
/// MIRR(values, finance_rate, reinvest_rate) - calculates the modified internal rate of return for a series of cash flows.
/// </summary>
public sealed class MirrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MirrFunction Instance = new();

    private MirrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MIRR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Last two arguments are finance_rate and reinvest_rate
        var financeRateArg = args[args.Length - 2];
        var reinvestRateArg = args[args.Length - 1];

        // Check for errors
        if (financeRateArg.IsError)
        {
            return financeRateArg;
        }

        if (reinvestRateArg.IsError)
        {
            return reinvestRateArg;
        }

        // Validate rates are numbers
        if (financeRateArg.Type != CellValueType.Number || reinvestRateArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var financeRate = financeRateArg.NumericValue;
        var reinvestRate = reinvestRateArg.NumericValue;

        // Extract cash flow values (all arguments except last two)
        var valueCount = args.Length - 2;
        var values = new double[valueCount];

        for (int i = 0; i < valueCount; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            values[i] = args[i].NumericValue;
        }

        // MIRR requires at least one positive and one negative cash flow
        bool hasPositive = false;
        bool hasNegative = false;

        foreach (var value in values)
        {
            if (value > 0)
            {
                hasPositive = true;
            }
            else if (value < 0)
            {
                hasNegative = true;
            }

            if (hasPositive && hasNegative)
            {
                break;
            }
        }

        if (!hasPositive || !hasNegative)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate present value of negative cash flows (costs) discounted at finance_rate
        double pvNegative = 0.0;
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i] < 0)
            {
                pvNegative += values[i] / System.Math.Pow(1 + financeRate, i);
            }
        }

        // Calculate future value of positive cash flows (returns) compounded at reinvest_rate
        double fvPositive = 0.0;
        int n = values.Length;
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i] > 0)
            {
                fvPositive += values[i] * System.Math.Pow(1 + reinvestRate, n - i - 1);
            }
        }

        // Check for division by zero
        if (pvNegative == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // MIRR formula: (FV_positive / -PV_negative)^(1/(n-1)) - 1
        var mirr = System.Math.Pow(-fvPositive / pvNegative, 1.0 / (n - 1)) - 1;

        if (double.IsNaN(mirr) || double.IsInfinity(mirr))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(mirr);
    }
}
