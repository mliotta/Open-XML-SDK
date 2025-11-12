// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CRITBINOM function (legacy compatibility function).
/// CRITBINOM(trials, probability_s, alpha) - returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
/// This is a legacy function that delegates to BINOM.INV.
/// </summary>
public sealed class CritbinomFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CritbinomFunction Instance = new();

    private CritbinomFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CRITBINOM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to BINOM.INV
        return BinomInvFunction.Instance.Execute(context, args);
    }
}
