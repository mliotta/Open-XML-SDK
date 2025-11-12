// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VAR.S function.
/// VAR.S(number1, [number2], ...) - variance (sample) - Excel 2010+ compatibility function.
/// This is the same as VAR.
/// </summary>
public sealed class VarSFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly VarSFunction Instance = new();

    private VarSFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VAR.S";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to VAR implementation
        return VarFunction.Instance.Execute(context, args);
    }
}
