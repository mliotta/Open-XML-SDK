// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEIBULL function (legacy version).
/// WEIBULL(x, alpha, beta, cumulative) - returns the Weibull distribution.
/// </summary>
public sealed class WeibullFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WeibullFunction Instance = new();

    private WeibullFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEIBULL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        return WeibullDistFunction.Instance.Execute(context, args);
    }
}
