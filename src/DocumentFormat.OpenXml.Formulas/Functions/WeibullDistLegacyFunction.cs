// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEIBULL function (legacy Excel 2007 compatibility).
/// WEIBULL(x, alpha, beta, cumulative) - returns the Weibull distribution.
/// This is the legacy version; modern Excel uses WEIBULL.DIST.
/// </summary>
public sealed class WeibullDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WeibullDistLegacyFunction Instance = new();

    private WeibullDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEIBULL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // WEIBULL has the same signature as WEIBULL.DIST
        // Delegate directly to WEIBULL.DIST
        return WeibullDistFunction.Instance.Execute(context, args);
    }
}
