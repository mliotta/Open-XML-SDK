// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FORMULATEXT function.
/// FORMULATEXT(reference) - Returns the formula at the given reference as text.
/// </summary>
public sealed class FormulaTextFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FormulaTextFunction Instance = new();

    private FormulaTextFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FORMULATEXT";

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
            return CellValue.Error("#VALUE!");
        }

        var cellReference = reference.StringValue;

        // Get the cell from the context
        // Note: This is a limitation - we need access to the actual Cell object, not just its value
        // For now, we'll return #N/A as this requires deeper integration with the worksheet
        return CellValue.Error("#N/A");
    }
}
