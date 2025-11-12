// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHIINV function (legacy Excel 2007 compatibility).
/// CHIINV(probability, degrees_freedom) - returns the inverse of the right-tailed probability of the chi-squared distribution.
/// This is the legacy version; modern Excel uses CHISQ.INV.RT.
/// </summary>
public sealed class ChiInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiInvLegacyFunction Instance = new();

    private ChiInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHIINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // CHIINV is equivalent to CHISQ.INV.RT
        return ChiSqInvRTFunction.Instance.Execute(context, args);
    }
}
