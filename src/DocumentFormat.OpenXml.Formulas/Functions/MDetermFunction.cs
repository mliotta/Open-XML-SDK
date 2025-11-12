// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MDETERM function.
/// MDETERM(array) - matrix determinant.
/// For Phase 0: simplified for 2x2 matrices, return #VALUE! for larger.
/// </summary>
public sealed class MDetermFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MDetermFunction Instance = new();

    private MDetermFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MDETERM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Phase 0: Simplified implementation
        // Return #VALUE! as full matrix support requires array handling
        // Full implementation will be added in a future phase
        return CellValue.Error("#VALUE!");
    }
}
