// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMADIST function (legacy Excel 2007 compatibility).
/// GAMMADIST(x, alpha, beta, cumulative) - returns the gamma distribution.
/// This is the legacy version; modern Excel uses GAMMA.DIST.
/// </summary>
public sealed class GammaDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammaDistLegacyFunction Instance = new();

    private GammaDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMADIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // GAMMADIST has the same signature as GAMMA.DIST
        // Delegate directly to GAMMA.DIST
        return GammaDistFunction.Instance.Execute(context, args);
    }
}
