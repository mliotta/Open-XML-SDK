// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the UNICODE function.
/// UNICODE(text) - returns Unicode code point for first character.
/// </summary>
public sealed class UnicodeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly UnicodeFunction Instance = new();

    private UnicodeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "UNICODE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        var text = args[0].StringValue;

        if (string.IsNullOrEmpty(text))
        {
            return CellValue.Error("#VALUE!");
        }

        // Get the code point of the first character
        // Handle surrogate pairs properly
        int codePoint;
        if (char.IsHighSurrogate(text[0]) && text.Length > 1 && char.IsLowSurrogate(text[1]))
        {
            // Surrogate pair - convert to code point
            codePoint = char.ConvertToUtf32(text[0], text[1]);
        }
        else
        {
            // Single character
            codePoint = text[0];
        }

        return CellValue.FromNumber(codePoint);
    }
}
