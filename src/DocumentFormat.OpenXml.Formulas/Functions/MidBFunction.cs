// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MIDB function.
/// MIDB(text, start_num, num_bytes) - returns substring based on byte count (UTF-8, 1-based indexing).
/// </summary>
public sealed class MidBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MidBFunction Instance = new();

    private MidBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MIDB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        var text = args[0].StringValue;

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

        // Get bytes from text
        var bytes = Encoding.UTF8.GetBytes(text);

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= bytes.Length || numBytes == 0)
        {
            return CellValue.FromString(string.Empty);
        }

        // Calculate the actual number of bytes to extract
        var length = System.Math.Min(numBytes, bytes.Length - startIndex);

        // Extract the bytes
        var resultBytes = new byte[length];
        System.Array.Copy(bytes, startIndex, resultBytes, 0, length);

        // Convert back to string, handling partial UTF-8 sequences
        var result = Encoding.UTF8.GetString(resultBytes, 0, length);

        // Clean up incomplete characters at both ends
        result = CleanupIncompleteCharacters(result);

        return CellValue.FromString(result);
    }

    private static string CleanupIncompleteCharacters(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        // Remove incomplete characters at the start
        while (text.Length > 0 && text[0] == '\uFFFD')
        {
            text = text.Substring(1);
        }

        // Remove incomplete characters at the end
        while (text.Length > 0 && text[text.Length - 1] == '\uFFFD')
        {
            text = text.Substring(0, text.Length - 1);
        }

        return text;
    }
}
