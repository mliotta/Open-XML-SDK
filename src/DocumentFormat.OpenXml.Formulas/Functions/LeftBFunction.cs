// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LEFTB function.
/// LEFTB(text, [num_bytes]) - returns leftmost characters based on byte count (UTF-8).
/// </summary>
public sealed class LeftBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LeftBFunction Instance = new();

    private LeftBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LEFTB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        var text = args[0].StringValue;
        var numBytes = 1;

        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            numBytes = (int)args[1].NumericValue;

            if (numBytes < 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        if (numBytes == 0)
        {
            return CellValue.FromString(string.Empty);
        }

        // Get bytes from text
        var bytes = Encoding.UTF8.GetBytes(text);

        if (numBytes >= bytes.Length)
        {
            return CellValue.FromString(text);
        }

        // Take only the requested number of bytes
        var resultBytes = new byte[numBytes];
        System.Array.Copy(bytes, 0, resultBytes, 0, numBytes);

        // Convert back to string, handling partial UTF-8 sequences
        var result = Encoding.UTF8.GetString(resultBytes, 0, numBytes);

        // Remove any incomplete characters at the end
        result = RemoveIncompleteCharacters(result);

        return CellValue.FromString(result);
    }

    private static string RemoveIncompleteCharacters(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        // Check if the last character is a replacement character (U+FFFD)
        // which indicates an incomplete UTF-8 sequence
        while (text.Length > 0 && text[text.Length - 1] == '\uFFFD')
        {
            text = text.Substring(0, text.Length - 1);
        }

        return text;
    }
}
