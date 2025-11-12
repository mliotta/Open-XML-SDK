// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHIDIST function (legacy Excel 2007 compatibility).
/// CHIDIST(x, degrees_freedom) - returns the right-tailed probability of the chi-squared distribution.
/// This is the legacy version; modern Excel uses CHISQ.DIST.RT.
/// </summary>
public sealed class ChiDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiDistLegacyFunction Instance = new();

    private ChiDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHIDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CHIDIST is equivalent to CHISQ.DIST.RT
        return ChiSqDistRTFunction.Instance.Execute(context, args);
    }
}
