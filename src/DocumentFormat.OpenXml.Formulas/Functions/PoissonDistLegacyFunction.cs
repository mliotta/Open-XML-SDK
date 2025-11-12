// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the POISSON function (legacy Excel 2007 compatibility).
/// POISSON(x, mean, cumulative) - returns the Poisson distribution.
/// This is the legacy version; modern Excel uses POISSON.DIST.
/// </summary>
public sealed class PoissonDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PoissonDistLegacyFunction Instance = new();

    private PoissonDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "POISSON";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // POISSON has the same signature as POISSON.DIST
        // Delegate directly to POISSON.DIST
        return PoissonDistFunction.Instance.Execute(context, args);
    }
}
