// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RRI function.
/// RRI(nper, pv, fv) - returns an equivalent interest rate for the growth of an investment.
/// </summary>
public sealed class RriFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RriFunction Instance = new();

    private RriFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RRI";

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

        var nper = args[0].NumericValue;
        var pv = args[1].NumericValue;
        var fv = args[2].NumericValue;

        // Validate arguments
        if (nper <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        if (pv == 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Formula: (fv/pv)^(1/nper) - 1
        var ratio = fv / pv;

        // Check for invalid ratio (can't take root of negative number with non-integer exponent)
        if (ratio < 0)
        {
            return CellValue.Error("#NUM!");
        }

        var rri = System.Math.Pow(ratio, 1.0 / nper) - 1;

        if (double.IsNaN(rri) || double.IsInfinity(rri))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(rri);
    }
}
