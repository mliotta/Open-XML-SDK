// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORMINV function (legacy Excel 2007 compatibility).
/// NORMINV(probability, mean, standard_dev) - returns the inverse of the normal cumulative distribution function.
/// This is the legacy version; modern Excel uses NORM.INV.
/// </summary>
public sealed class NormInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormInvLegacyFunction Instance = new();

    private NormInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORMINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // NORMINV has the same signature as NORM.INV
        // Delegate directly to NORM.INV
        return NormInvFunction.Instance.Execute(context, args);
    }
}
