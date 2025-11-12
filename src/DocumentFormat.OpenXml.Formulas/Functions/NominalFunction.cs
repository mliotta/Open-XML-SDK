// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NOMINAL function.
/// NOMINAL(effect_rate, npery) - calculates the nominal annual interest rate given the effective rate and number of compounding periods per year.
/// </summary>
public sealed class NominalFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NominalFunction Instance = new();

    private NominalFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NOMINAL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Validate arguments are numbers
        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var effectiveRate = args[0].NumericValue;
        var npery = args[1].NumericValue;

        // Validate inputs
        if (effectiveRate <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // npery must be >= 1 and will be truncated to integer
        var periodsPerYear = (int)System.Math.Truncate(npery);
        if (periodsPerYear < 1)
        {
            return CellValue.Error("#NUM!");
        }

        // NOMINAL formula: npery * ((1 + effect_rate)^(1/npery) - 1)
        var nominalRate = periodsPerYear * (System.Math.Pow(1 + effectiveRate, 1.0 / periodsPerYear) - 1);

        if (double.IsNaN(nominalRate) || double.IsInfinity(nominalRate))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(nominalRate);
    }
}
