// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CELL function.
/// CELL(info_type, [reference]) - Returns information about the formatting, location, or contents of a cell.
/// </summary>
/// <remarks>
/// This is a simplified implementation. Full implementation would require access to:
/// - Cell formatting information (number format, alignment, color, etc.)
/// - Cell location (address, row, column, sheet name)
/// - Cell protection status
/// - Workbook file path
///
/// Supported info_type values:
/// "address" - Cell reference as text
/// "col" - Column number
/// "color" - 1 if cell is formatted in color for negative values, 0 otherwise
/// "contents" - Value of the cell
/// "filename" - Filename and path
/// "format" - Number format code
/// "parentheses" - 1 if cell is formatted with parentheses, 0 otherwise
/// "prefix" - Text alignment prefix (', ", ^, or \)
/// "protect" - 1 if cell is locked, 0 otherwise
/// "row" - Row number
/// "type" - Type of data (b=blank, l=label/text, v=value)
/// "width" - Column width
/// </remarks>
public sealed class CellFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CellFunction Instance = new();

    private CellFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CELL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        var infoType = args[0].StringValue.ToLowerInvariant();

        // If no reference is provided, use the current cell (context)
        // For this simplified implementation, we'll return default values
        var reference = args.Length == 2 ? args[1] : CellValue.Empty;

        // Check for errors in reference
        if (reference.IsError)
        {
            return reference;
        }

        return infoType switch
        {
            "address" => CellValue.FromString("$A$1"),
            "col" => CellValue.FromNumber(1),
            "color" => CellValue.FromNumber(0),
            "contents" => reference.Type == CellValueType.Empty ? CellValue.FromString("") : reference,
            "filename" => CellValue.FromString(""),
            "format" => CellValue.FromString("G"),
            "parentheses" => CellValue.FromNumber(0),
            "prefix" => CellValue.FromString(""),
            "protect" => CellValue.FromNumber(1),
            "row" => CellValue.FromNumber(1),
            "type" => GetCellType(reference),
            "width" => CellValue.FromNumber(10),
            _ => CellValue.Error("#VALUE!"),
        };
    }

    private static CellValue GetCellType(CellValue value)
    {
        var typeChar = value.Type switch
        {
            CellValueType.Empty => "b",
            CellValueType.Text => "l",
            CellValueType.Number => "v",
            CellValueType.Boolean => "v",
            CellValueType.Error => "v",
            _ => "b",
        };

        return CellValue.FromString(typeChar);
    }
}
