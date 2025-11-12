// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the POISSON function (legacy version).
/// POISSON(x, mean, cumulative) - returns the Poisson distribution.
/// </summary>
public sealed class PoissonFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PoissonFunction Instance = new();

    private PoissonFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "POISSON";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        return PoissonDistFunction.Instance.Execute(context, args);
    }
}
