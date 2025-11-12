// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FINDB function.
/// FINDB(find_text, within_text, [start_num]) - finds text by byte position (case-sensitive, UTF-8, 1-based).
/// </summary>
public sealed class FindBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FindBFunction Instance = new();

    private FindBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FINDB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
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

        var findText = args[0].StringValue;
        var withinText = args[1].StringValue;
        var startNum = 1;

        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            startNum = (int)args[2].NumericValue;

            if (startNum < 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Convert to bytes
        var withinBytes = Encoding.UTF8.GetBytes(withinText);
        var findBytes = Encoding.UTF8.GetBytes(findText);

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= withinBytes.Length)
        {
            return CellValue.Error("#VALUE!");
        }

        // Find the byte position
        var position = FindBytePattern(withinBytes, findBytes, startIndex);

        if (position == -1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Return 1-based position
        return CellValue.FromNumber(position + 1);
    }

    private static int FindBytePattern(byte[] haystack, byte[] needle, int startIndex)
    {
        if (needle.Length == 0)
        {
            return startIndex;
        }

        for (var i = startIndex; i <= haystack.Length - needle.Length; i++)
        {
            var found = true;
            for (var j = 0; j < needle.Length; j++)
            {
                if (haystack[i + j] != needle[j])
                {
                    found = false;
                    break;
                }
            }

            if (found)
            {
                return i;
            }
        }

        return -1;
    }
}
