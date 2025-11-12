// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERCENTRANK.INC function.
/// PERCENTRANK.INC(array, x, [significance]) - returns the rank of a value as a percentage (0 to 1 inclusive).
/// </summary>
public sealed class PercentrankIncFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PercentrankIncFunction Instance = new();

    private PercentrankIncFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERCENTRANK.INC";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // PERCENTRANK.INC is the same as PERCENTRANK
        return PercentrankFunction.Instance.Execute(context, args);
    }
}
