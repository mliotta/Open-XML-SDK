// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MODE.SNGL function.
/// MODE.SNGL(number1, [number2], ...) - returns most frequent value (single mode) - Excel 2010+ compatibility function.
/// This is the same as MODE.
/// </summary>
public sealed class ModeSnglFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ModeSnglFunction Instance = new();

    private ModeSnglFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MODE.SNGL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to MODE implementation
        return ModeFunction.Instance.Execute(context, args);
    }
}
