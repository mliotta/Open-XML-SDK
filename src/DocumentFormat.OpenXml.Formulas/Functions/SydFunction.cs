// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SYD function.
/// SYD(cost, salvage, life, period) - calculates depreciation using sum-of-years' digits method.
/// </summary>
public sealed class SydFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SydFunction Instance = new();

    private SydFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SYD";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        for (int i = 0; i < args.Length; i++)
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

        var cost = args[0].NumericValue;
        var salvage = args[1].NumericValue;
        var life = args[2].NumericValue;
        var period = args[3].NumericValue;

        // Validate inputs
        if (life <= 0 || period < 1 || period > life)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate sum-of-years' digits depreciation
        // Formula: ((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1))
        var depreciableAmount = cost - salvage;
        var sumOfYears = life * (life + 1);
        var yearDigit = life - period + 1;

        var depreciation = (depreciableAmount * yearDigit * 2.0) / sumOfYears;

        if (double.IsNaN(depreciation) || double.IsInfinity(depreciation))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(depreciation);
    }
}
