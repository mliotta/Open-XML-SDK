// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORMSINV function (legacy Excel 2007 compatibility).
/// NORMSINV(probability) - returns the inverse of the standard normal cumulative distribution function.
/// This is the legacy version; modern Excel uses NORM.S.INV.
/// </summary>
public sealed class NormSInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormSInvLegacyFunction Instance = new();

    private NormSInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORMSINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // NORMSINV has the same signature as NORM.S.INV
        // Delegate directly to NORM.S.INV
        return NormSInvFunction.Instance.Execute(context, args);
    }
}
