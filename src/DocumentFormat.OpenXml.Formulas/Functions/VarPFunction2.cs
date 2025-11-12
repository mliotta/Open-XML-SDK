// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VAR.P function.
/// VAR.P(number1, [number2], ...) - variance (population) - Excel 2010+ compatibility function.
/// This is the same as VARP.
/// </summary>
public sealed class VarPFunction2 : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly VarPFunction2 Instance = new();

    private VarPFunction2()
    {
    }

    /// <inheritdoc/>
    public string Name => "VAR.P";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // Delegate to VARP implementation
        return VarPFunction.Instance.Execute(context, args);
    }
}
