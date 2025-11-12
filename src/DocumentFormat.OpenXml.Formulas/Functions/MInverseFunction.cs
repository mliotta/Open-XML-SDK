// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MINVERSE function.
/// MINVERSE(array) - matrix inverse.
/// For Phase 0: simplified for 2x2 matrices, return first element.
/// </summary>
public sealed class MInverseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MInverseFunction Instance = new();

    private MInverseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MINVERSE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Phase 0: Simplified implementation
        // Return #VALUE! as full matrix support requires array handling
        // Full implementation will be added in a future phase
        return CellValue.Error("#VALUE!");
    }
}
