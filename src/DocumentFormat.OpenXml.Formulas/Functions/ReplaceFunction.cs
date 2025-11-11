// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REPLACE function.
/// REPLACE(old_text, start_num, num_chars, new_text) - replaces part of text string.
/// </summary>
public sealed class ReplaceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ReplaceFunction Instance = new();

    private ReplaceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "REPLACE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[2].IsError)
        {
            return args[2];
        }

        if (args[3].IsError)
        {
            return args[3];
        }

        var oldText = args[0].StringValue;

        if (args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var startNum = (int)args[1].NumericValue;
        var numChars = (int)args[2].NumericValue;
        var newText = args[3].StringValue;

        if (startNum < 1 || numChars < 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        // If start position is beyond the text length, append new text
        if (startIndex >= oldText.Length)
        {
            return CellValue.FromString(oldText + newText);
        }

        // Calculate the end position of the replacement
        var endIndex = System.Math.Min(startIndex + numChars, oldText.Length);

        // Build the result: part before + new text + part after
        var result = oldText.Substring(0, startIndex) + newText + oldText.Substring(endIndex);

        return CellValue.FromString(result);
    }
}
