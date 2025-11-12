// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ANCHORARRAY function.
/// ANCHORARRAY(reference)
/// Returns the anchor array reference for a dynamic array formula.
///
/// Phase 0 Implementation:
/// - Returns the reference value itself (simplified)
/// - Full dynamic array anchor detection requires engine enhancements
/// </summary>
public sealed class AnchorArrayFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AnchorArrayFunction Instance = new AnchorArrayFunction();

    private AnchorArrayFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ANCHORARRAY";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors
        if (args[0].IsError)
        {
            return args[0];
        }

        // Phase 0: Simply return the first argument
        // In a full implementation, this would analyze the reference
        // and return the top-left cell of the spilled array range
        return args[0];
    }
}
