// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EXPONDIST function (legacy Excel 2007 compatibility).
/// EXPONDIST(x, lambda, cumulative) - returns the exponential distribution.
/// This is the legacy version; modern Excel uses EXPON.DIST.
/// </summary>
public sealed class ExponDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ExponDistLegacyFunction Instance = new();

    private ExponDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EXPONDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // EXPONDIST has the same signature as EXPON.DIST
        // Delegate directly to EXPON.DIST
        return ExponDistFunction.Instance.Execute(context, args);
    }
}
