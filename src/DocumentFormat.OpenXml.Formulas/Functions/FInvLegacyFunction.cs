// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FINV function (legacy Excel 2007 compatibility).
/// FINV(probability, degrees_freedom1, degrees_freedom2) - returns the inverse of the right-tailed F probability distribution.
/// This is the legacy version; modern Excel uses F.INV.RT.
/// </summary>
public sealed class FInvLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FInvLegacyFunction Instance = new();

    private FInvLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FINV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // FINV is equivalent to F.INV.RT
        return FInvRTFunction.Instance.Execute(context, args);
    }
}
