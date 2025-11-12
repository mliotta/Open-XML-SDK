// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TEXTJOIN function.
/// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...) - joins text with delimiter.
/// </summary>
public sealed class TextJoinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TextJoinFunction Instance = new();

    private TextJoinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TEXTJOIN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument: delimiter
        if (args[0].IsError)
        {
            return args[0];
        }

        var delimiter = args[0].StringValue;

        // Second argument: ignore_empty (boolean)
        if (args[1].IsError)
        {
            return args[1];
        }

        bool ignoreEmpty;
        if (args[1].Type == CellValueType.Boolean)
        {
            ignoreEmpty = args[1].BoolValue;
        }
        else if (args[1].Type == CellValueType.Number)
        {
            // Excel allows numeric values: 0 = FALSE, non-zero = TRUE
            ignoreEmpty = args[1].NumericValue != 0;
        }
        else
        {
            // Try to parse text as boolean
            var text = args[1].StringValue.ToUpperInvariant();
            if (text == "TRUE")
            {
                ignoreEmpty = true;
            }
            else if (text == "FALSE")
            {
                ignoreEmpty = false;
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Remaining arguments: text values to join
        var result = new StringBuilder();
        var first = true;

        for (int i = 2; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i]; // Propagate errors
            }

            var text = args[i].StringValue;

            // Skip empty strings if ignore_empty is TRUE
            if (ignoreEmpty && string.IsNullOrEmpty(text))
            {
                continue;
            }

            if (!first)
            {
                result.Append(delimiter);
            }

            result.Append(text);
            first = false;
        }

        return CellValue.FromString(result.ToString());
    }
}
