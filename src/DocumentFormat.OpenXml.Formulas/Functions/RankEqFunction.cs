// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RANK.EQ function.
/// RANK.EQ(number, ref, [order]) - returns rank of number in list (same as RANK).
/// </summary>
public sealed class RankEqFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RankEqFunction Instance = new();

    private RankEqFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RANK.EQ";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // RANK.EQ is the same as RANK - returns the same rank for duplicate values
        return RankFunction.Instance.Execute(context, args);
    }
}
