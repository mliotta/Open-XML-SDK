// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PDURATION function.
/// PDURATION(rate, pv, fv) - returns the number of periods required for an investment to reach a specified value.
/// </summary>
public sealed class PdurationFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PdurationFunction Instance = new();

    private PdurationFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PDURATION";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in all arguments
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

        var rate = args[0].NumericValue;
        var pv = args[1].NumericValue;
        var fv = args[2].NumericValue;

        // Validate arguments
        if (rate <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        if (pv <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        if (fv <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Formula: LOG(fv/pv) / LOG(1+rate)
        var pduration = System.Math.Log(fv / pv) / System.Math.Log(1 + rate);

        if (double.IsNaN(pduration) || double.IsInfinity(pduration))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(pduration);
    }
}
