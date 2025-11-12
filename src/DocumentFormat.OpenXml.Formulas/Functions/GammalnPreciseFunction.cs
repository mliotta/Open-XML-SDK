// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMALN.PRECISE function.
/// GAMMALN.PRECISE(x) - returns the natural logarithm of the gamma function (same as GAMMALN).
/// </summary>
public sealed class GammalnPreciseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammalnPreciseFunction Instance = new();

    private GammalnPreciseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMALN.PRECISE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        return GammalnFunction.Instance.Execute(context, args);
    }
}
