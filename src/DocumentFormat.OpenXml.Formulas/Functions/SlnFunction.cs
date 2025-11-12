// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SLN function.
/// SLN(cost, salvage, life) - calculates straight-line depreciation for one period.
/// </summary>
public sealed class SlnFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SlnFunction Instance = new();

    private SlnFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SLN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        // Validate life is positive
        if (life <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate straight-line depreciation
        var depreciation = (cost - salvage) / life;

        if (double.IsNaN(depreciation) || double.IsInfinity(depreciation))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(depreciation);
    }
}
