// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PEARSON function.
/// PEARSON(array1, array2) - returns the Pearson product-moment correlation coefficient.
/// This function is identical to CORREL.
/// </summary>
public sealed class PearsonFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PearsonFunction Instance = new();

    private PearsonFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PEARSON";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // PEARSON is identical to CORREL
        return CorrelFunction.Instance.Execute(context, args);
    }
}
