// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SHEET function.
/// SHEET([value]) - Returns the sheet number of the referenced sheet.
/// If no argument is provided, returns the sheet number of the sheet containing the formula.
/// </summary>
public sealed class SheetFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SheetFunction Instance = new();

    private SheetFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SHEET";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            // Return the current sheet number (1-indexed)
            // Note: This requires context to know which sheet we're on
            // For now, we'll return 1 as the default
            return CellValue.FromNumber(1);
        }

        var reference = args[0];

        // Check for errors
        if (reference.IsError)
        {
            return reference;
        }

        // If a reference is provided, we would need to parse it and determine the sheet
        // For now, return 1 as a default
        return CellValue.FromNumber(1);
    }
}
