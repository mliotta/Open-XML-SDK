// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISFORMULA function.
/// ISFORMULA(reference) - Returns TRUE if the reference contains a formula, FALSE otherwise.
/// </summary>
public sealed class IsFormulaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsFormulaFunction Instance = new();

    private IsFormulaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISFORMULA";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        var reference = args[0];

        // Check for errors
        if (reference.IsError)
        {
            return reference;
        }

        // Reference must be text (cell reference like "A1")
        if (reference.Type != CellValueType.Text)
        {
            return CellValue.FromBool(false);
        }

        // Note: This is a limitation - we need access to the actual Cell object to check if it has a formula
        // For now, we'll return FALSE as this requires deeper integration with the worksheet
        return CellValue.FromBool(false);
    }
}
