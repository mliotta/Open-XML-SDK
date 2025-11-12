// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISPMT function.
/// ISPMT(rate, per, nper, pv) - calculates the interest paid during a specific period of an investment.
/// This is for straight-line depreciation of principal (different from IPMT).
/// </summary>
public sealed class IsPmtFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsPmtFunction Instance = new();

    private IsPmtFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISPMT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in all arguments
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

        var rate = args[0].NumericValue;
        var per = args[1].NumericValue;
        var nper = args[2].NumericValue;
        var pv = args[3].NumericValue;

        // Validate nper is not zero
        if (nper == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Formula: -pv * rate * (per/nper - 1)
        // This represents straight-line principal reduction
        var ispmt = -pv * rate * (per / nper - 1);

        if (double.IsNaN(ispmt) || double.IsInfinity(ispmt))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(ispmt);
    }
}
