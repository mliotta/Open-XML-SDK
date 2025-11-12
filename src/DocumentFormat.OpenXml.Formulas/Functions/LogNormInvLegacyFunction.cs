// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LOGINV function (legacy Excel 2007 compatibility).
/// LOGINV(probability, mean, standard_dev) - returns the inverse of the lognormal cumulative distribution function.
/// This is the legacy version; modern Excel uses LOGNORM.INV.
/// </summary>
public sealed class LogNormInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LogNormInvLegacyFunction Instance = new();

    private LogNormInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LOGINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // LOGINV has the same signature as LOGNORM.INV
        // Delegate directly to LOGNORM.INV
        return LogNormInvFunction.Instance.Execute(context, args);
    }
}
