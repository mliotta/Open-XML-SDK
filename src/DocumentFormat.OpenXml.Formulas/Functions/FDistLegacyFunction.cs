// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FDIST function (legacy Excel 2007 compatibility).
/// FDIST(x, degrees_freedom1, degrees_freedom2) - returns the right-tailed F probability distribution.
/// This is the legacy version; modern Excel uses F.DIST.RT.
/// </summary>
public sealed class FDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FDistLegacyFunction Instance = new();

    private FDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FDIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // FDIST is equivalent to F.DIST.RT
        return FDistRTFunction.Instance.Execute(context, args);
    }
}
