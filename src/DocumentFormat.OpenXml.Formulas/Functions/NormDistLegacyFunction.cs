// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORMDIST function (legacy Excel 2007 compatibility).
/// NORMDIST(x, mean, standard_dev, cumulative) - returns the normal distribution.
/// This is the legacy version; modern Excel uses NORM.DIST.
/// </summary>
public sealed class NormDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormDistLegacyFunction Instance = new();

    private NormDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORMDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // NORMDIST has the same signature as NORM.DIST
        // Delegate directly to NORM.DIST
        return NormDistFunction.Instance.Execute(context, args);
    }
}
