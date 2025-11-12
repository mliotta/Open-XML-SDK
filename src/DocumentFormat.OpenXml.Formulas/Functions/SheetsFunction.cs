// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SHEETS function.
/// SHEETS([reference]) - Returns the number of sheets in a reference.
/// If no argument is provided, returns the number of sheets in the workbook.
/// </summary>
public sealed class SheetsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SheetsFunction Instance = new();

    private SheetsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SHEETS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            // Return the total number of sheets in the workbook
            // Note: This requires access to the workbook part
            // For now, we'll return 1 as the default
            return CellValue.FromNumber(1);
        }

        var reference = args[0];

        // Check for errors
        if (reference.IsError)
        {
            return reference;
        }

        // If a reference is provided, we would need to parse it and count the sheets
        // For now, return 1 as a default
        return CellValue.FromNumber(1);
    }
}
