// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EFFECT function.
/// EFFECT(nominal_rate, npery) - calculates the effective annual interest rate given the nominal rate and number of compounding periods per year.
/// </summary>
public sealed class EffectFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly EffectFunction Instance = new();

    private EffectFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EFFECT";

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

        var nominalRate = args[0].NumericValue;
        var npery = args[1].NumericValue;

        // Validate inputs
        if (nominalRate <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // npery must be >= 1 and will be truncated to integer
        var periodsPerYear = (int)System.Math.Truncate(npery);
        if (periodsPerYear < 1)
        {
            return CellValue.Error("#NUM!");
        }

        // EFFECT formula: (1 + nominal_rate/npery)^npery - 1
        var effectiveRate = System.Math.Pow(1 + nominalRate / periodsPerYear, periodsPerYear) - 1;

        if (double.IsNaN(effectiveRate) || double.IsInfinity(effectiveRate))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(effectiveRate);
    }
}
