// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TEXTSPLIT function.
/// TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]) - splits text into array.
/// </summary>
public sealed class TextSplitFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TextSplitFunction Instance = new();

    private TextSplitFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TEXTSPLIT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var text = args[0].StringValue;
        var colDelimiter = args[1].StringValue;

        // Default values
        string? rowDelimiter = null;
        var ignoreEmpty = false;
        var matchMode = 0; // 0 = case-sensitive, 1 = case-insensitive
        var padWith = string.Empty;

        // Parse optional arguments
        if (args.Length >= 3 && args[2].Type != CellValueType.Empty)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            rowDelimiter = args[2].StringValue;
        }

        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Boolean)
            {
                ignoreEmpty = args[3].BoolValue;
            }
            else if (args[3].Type == CellValueType.Number)
            {
                ignoreEmpty = args[3].NumericValue != 0;
            }
        }

        if (args.Length >= 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == CellValueType.Number)
            {
                matchMode = (int)args[4].NumericValue;
                if (matchMode != 0 && matchMode != 1)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        if (args.Length >= 6)
        {
            if (args[5].IsError)
            {
                return args[5];
            }

            padWith = args[5].StringValue;
        }

        // Empty text returns single cell with empty string
        if (string.IsNullOrEmpty(text))
        {
            return CellValue.FromString(string.Empty);
        }

        var comparisonType = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        // Split by row delimiter first (if provided)
        var rows = new List<string>();
        if (!string.IsNullOrEmpty(rowDelimiter))
        {
            rows.AddRange(SplitString(text, rowDelimiter, comparisonType));
        }
        else
        {
            rows.Add(text);
        }

        // Split each row by column delimiter
        var resultGrid = new List<List<string>>();
        var maxCols = 0;

        foreach (var row in rows)
        {
            var cols = new List<string>();
            if (string.IsNullOrEmpty(colDelimiter))
            {
                // No column delimiter means single column
                cols.Add(row);
            }
            else
            {
                var splitCols = SplitString(row, colDelimiter, comparisonType);
                if (ignoreEmpty)
                {
                    foreach (var col in splitCols)
                    {
                        if (!string.IsNullOrEmpty(col))
                        {
                            cols.Add(col);
                        }
                    }
                }
                else
                {
                    cols.AddRange(splitCols);
                }
            }

            if (cols.Count > maxCols)
            {
                maxCols = cols.Count;
            }

            resultGrid.Add(cols);
        }

        // Pad rows to same length if needed
        foreach (var row in resultGrid)
        {
            while (row.Count < maxCols)
            {
                row.Add(padWith);
            }
        }

        // For now, return as concatenated text (full array support would require array value type)
        // Return first row, first column as simplified implementation
        if (resultGrid.Count > 0 && resultGrid[0].Count > 0)
        {
            return CellValue.FromString(resultGrid[0][0]);
        }

        return CellValue.FromString(string.Empty);
    }

    private static List<string> SplitString(string text, string delimiter, StringComparison comparison)
    {
        var result = new List<string>();
        var currentIndex = 0;

        while (currentIndex < text.Length)
        {
            var position = text.IndexOf(delimiter, currentIndex, comparison);
            if (position == -1)
            {
                result.Add(text.Substring(currentIndex));
                break;
            }

            result.Add(text.Substring(currentIndex, position - currentIndex));
            currentIndex = position + delimiter.Length;
        }

        // Handle trailing delimiter
        if (currentIndex == text.Length && text.EndsWith(delimiter, comparison))
        {
            result.Add(string.Empty);
        }

        return result;
    }
}
