// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BETAINV function (legacy Excel 2007 compatibility).
/// BETAINV(probability, alpha, beta, [A], [B]) - returns the inverse of the beta cumulative distribution function.
/// This is the legacy version; modern Excel uses BETA.INV.
/// </summary>
public sealed class BetaInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BetaInvLegacyFunction Instance = new();

    private BetaInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BETAINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // BETAINV has the same signature as BETA.INV
        // Delegate directly to BETA.INV
        return BetaInvFunction.Instance.Execute(context, args);
    }
}
