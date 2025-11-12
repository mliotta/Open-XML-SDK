// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONFIDENCE.NORM function.
/// CONFIDENCE.NORM(alpha, standard_dev, size) - returns the confidence interval for a population mean (normal distribution).
/// This is the same as CONFIDENCE.
/// </summary>
public sealed class ConfidenceNormFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConfidenceNormFunction Instance = new();

    private ConfidenceNormFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONFIDENCE.NORM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to CONFIDENCE implementation
        return ConfidenceFunction.Instance.Execute(context, args);
    }
}
