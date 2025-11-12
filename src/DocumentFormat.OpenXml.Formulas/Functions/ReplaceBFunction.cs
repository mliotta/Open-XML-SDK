// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REPLACEB function.
/// REPLACEB(old_text, start_num, num_bytes, new_text) - replaces text based on byte position (UTF-8, 1-based).
/// </summary>
public sealed class ReplaceBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ReplaceBFunction Instance = new();

    private ReplaceBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "REPLACEB";

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
        var newText = args[3].StringValue;

        if (args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var startNum = (int)args[1].NumericValue;
        var numBytes = (int)args[2].NumericValue;

        if (startNum < 1 || numBytes < 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get bytes from old text
        var oldBytes = Encoding.UTF8.GetBytes(oldText);

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        // Handle case where start position is beyond text length
        if (startIndex > oldBytes.Length)
        {
            // Just append the new text
            return CellValue.FromString(oldText + newText);
        }

        // Calculate end position
        var endIndex = System.Math.Min(startIndex + numBytes, oldBytes.Length);
        var bytesToRemove = endIndex - startIndex;

        // Build result: before + new + after
        var beforeBytes = new byte[startIndex];
        System.Array.Copy(oldBytes, 0, beforeBytes, 0, startIndex);

        var afterStart = startIndex + bytesToRemove;
        var afterLength = oldBytes.Length - afterStart;
        var afterBytes = new byte[afterLength];
        if (afterLength > 0)
        {
            System.Array.Copy(oldBytes, afterStart, afterBytes, 0, afterLength);
        }

        var newBytes = Encoding.UTF8.GetBytes(newText);

        // Combine all parts
        var resultBytes = new byte[beforeBytes.Length + newBytes.Length + afterBytes.Length];
        System.Array.Copy(beforeBytes, 0, resultBytes, 0, beforeBytes.Length);
        System.Array.Copy(newBytes, 0, resultBytes, beforeBytes.Length, newBytes.Length);
        System.Array.Copy(afterBytes, 0, resultBytes, beforeBytes.Length + newBytes.Length, afterBytes.Length);

        var result = Encoding.UTF8.GetString(resultBytes);

        return CellValue.FromString(result);
    }
}
