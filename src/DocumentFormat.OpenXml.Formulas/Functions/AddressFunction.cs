// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ADDRESS function.
/// ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text]) - Creates a cell reference as text.
/// </summary>
public sealed class AddressFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly AddressFunction Instance = new();

    private AddressFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ADDRESS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract row_num (required)
        var rowNumArg = args[0];
        if (rowNumArg.IsError)
        {
            return rowNumArg;
        }

        if (rowNumArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rowNum = (int)rowNumArg.NumericValue;
        if (rowNum < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract column_num (required)
        var colNumArg = args[1];
        if (colNumArg.IsError)
        {
            return colNumArg;
        }

        if (colNumArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var colNum = (int)colNumArg.NumericValue;
        if (colNum < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract abs_num (optional, default = 1 for absolute)
        var absNum = 1;
        if (args.Length >= 3)
        {
            var absNumArg = args[2];
            if (absNumArg.IsError)
            {
                return absNumArg;
            }

            if (absNumArg.Type == CellValueType.Number)
            {
                absNum = (int)absNumArg.NumericValue;
                if (absNum < 1 || absNum > 4)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        // Extract a1 (optional, default = TRUE for A1 notation)
        var useA1 = true;
        if (args.Length >= 4)
        {
            var a1Arg = args[3];
            if (a1Arg.IsError)
            {
                return a1Arg;
            }

            if (a1Arg.Type == CellValueType.Boolean)
            {
                useA1 = a1Arg.BoolValue;
            }
            else if (a1Arg.Type == CellValueType.Number)
            {
                useA1 = a1Arg.NumericValue != 0;
            }
        }

        // Extract sheet_text (optional)
        string? sheetText = null;
        if (args.Length >= 5)
        {
            var sheetArg = args[4];
            if (sheetArg.IsError)
            {
                return sheetArg;
            }

            if (sheetArg.Type == CellValueType.Text)
            {
                sheetText = sheetArg.StringValue;
            }
        }

        // Build the address
        string address;
        if (useA1)
        {
            // A1 notation
            address = BuildA1Address(rowNum, colNum, absNum);
        }
        else
        {
            // R1C1 notation
            address = BuildR1C1Address(rowNum, colNum, absNum);
        }

        // Prepend sheet name if provided
        if (!string.IsNullOrEmpty(sheetText))
        {
            // Sheet names with spaces or special characters should be quoted
            if (sheetText.Contains(" ") || sheetText.Contains("!"))
            {
                address = $"'{sheetText}'!{address}";
            }
            else
            {
                address = $"{sheetText}!{address}";
            }
        }

        return CellValue.FromString(address);
    }

    private static string BuildA1Address(int row, int col, int absNum)
    {
        var colLetter = GetColumnLetter(col);
        var rowStr = row.ToString(System.Globalization.CultureInfo.InvariantCulture);

        return absNum switch
        {
            1 => $"${colLetter}${rowStr}", // Absolute row and column
            2 => $"${colLetter}{rowStr}",  // Absolute column, relative row
            3 => $"{colLetter}${rowStr}",  // Relative column, absolute row
            4 => $"{colLetter}{rowStr}",   // Relative row and column
            _ => $"${colLetter}${rowStr}", // Default to absolute
        };
    }

    private static string BuildR1C1Address(int row, int col, int absNum)
    {
        var rowPart = absNum switch
        {
            1 or 2 => $"R{row}",     // Absolute row
            3 or 4 => $"R[{row}]",   // Relative row (offset notation)
            _ => $"R{row}",
        };

        var colPart = absNum switch
        {
            1 or 3 => $"C{col}",     // Absolute column
            2 or 4 => $"C[{col}]",   // Relative column (offset notation)
            _ => $"C{col}",
        };

        return $"{rowPart}{colPart}";
    }

    private static string GetColumnLetter(int column)
    {
        var result = string.Empty;

        while (column > 0)
        {
            var modulo = (column - 1) % 26;
            result = (char)('A' + modulo) + result;
            column = (column - modulo) / 26;
        }

        return result;
    }
}
