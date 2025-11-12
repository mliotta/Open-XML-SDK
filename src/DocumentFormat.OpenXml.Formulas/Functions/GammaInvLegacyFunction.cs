// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMAINV function (legacy Excel 2007 compatibility).
/// GAMMAINV(probability, alpha, beta) - returns the inverse of the gamma cumulative distribution function.
/// This is the legacy version; modern Excel uses GAMMA.INV.
/// </summary>
public sealed class GammaInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammaInvLegacyFunction Instance = new();

    private GammaInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMAINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // GAMMAINV has the same signature as GAMMA.INV
        // Delegate directly to GAMMA.INV
        return GammaInvFunction.Instance.Execute(context, args);
    }
}
