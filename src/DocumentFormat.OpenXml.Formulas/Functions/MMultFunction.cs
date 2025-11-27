// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MMULT function.
/// MMULT(array1, array2) - matrix multiplication.
/// For Phase 0: simplified, multiply corresponding elements.
/// </summary>
public sealed class MMultFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MMultFunction Instance = new();

    private MMultFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MMULT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Phase 0: Simplified implementation
        // If both are numbers, multiply them
        if (args[0].Type == FormulaResultType.Number && args[1].Type == FormulaResultType.Number)
        {
            return FormulaResult.FromNumber(args[0].NumericValue * args[1].NumericValue);
        }

        // Full matrix multiplication requires array support
        return FormulaResult.Error("#VALUE!");
    }
}
