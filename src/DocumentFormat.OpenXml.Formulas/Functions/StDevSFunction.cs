// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the STDEV.S function.
/// STDEV.S(number1, [number2], ...) - standard deviation (sample) - Excel 2010+ compatibility function.
/// This is the same as STDEV.
/// </summary>
public sealed class StDevSFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly StDevSFunction Instance = new();

    private StDevSFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "STDEV.S";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to STDEV implementation
        return StDevFunction.Instance.Execute(context, args);
    }
}
