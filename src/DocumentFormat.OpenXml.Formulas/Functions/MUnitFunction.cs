// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MUNIT function.
/// MUNIT(dimension) - identity matrix.
/// For Phase 0: return 1 (first element of identity matrix).
/// </summary>
public sealed class MUnitFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MUnitFunction Instance = new();

    private MUnitFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MUNIT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var dimension = args[0].NumericValue;

        // Dimension must be positive
        if (dimension <= 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Truncate to integer
        var n = (int)System.Math.Floor(dimension);

        // Phase 0: return 1 (first element of identity matrix)
        // In a full implementation, this would return an nxn identity matrix
        // where the diagonal is 1 and all other elements are 0
        return CellValue.FromNumber(1);
    }
}
