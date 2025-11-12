// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BINOMDIST function (legacy Excel 2007 compatibility).
/// BINOMDIST(number_s, trials, probability_s, cumulative) - returns the individual term binomial distribution probability.
/// This is the legacy version; modern Excel uses BINOM.DIST.
/// </summary>
public sealed class BinomDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BinomDistLegacyFunction Instance = new();

    private BinomDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BINOMDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // BINOMDIST has the same signature as BINOM.DIST
        // Delegate directly to BINOM.DIST
        return BinomDistFunction.Instance.Execute(context, args);
    }
}
