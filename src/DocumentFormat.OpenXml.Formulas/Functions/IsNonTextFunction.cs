// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISNONTEXT function.
/// ISNONTEXT(value) - TRUE if value is not text.
/// </summary>
public sealed class IsNonTextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsNonTextFunction Instance = new();

    private IsNonTextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISNONTEXT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isNonText = args[0].Type != CellValueType.Text;
        return CellValue.FromBool(isNonText);
    }
}
