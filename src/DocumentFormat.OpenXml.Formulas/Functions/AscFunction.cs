// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ASC function.
/// ASC(text) - Converts full-width (double-byte) characters to half-width (single-byte) characters.
/// Used primarily for Japanese text.
/// </summary>
public sealed class AscFunction : IFunctionImplementation
{
    public static readonly AscFunction Instance = new();

    private AscFunction()
    {
    }

    public string Name => "ASC";

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

        if (args[0].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        var text = args[0].StringValue;
        var result = new StringBuilder(text.Length);

        foreach (var ch in text)
        {
            // Full-width space to half-width space
            if (ch == '\u3000')
            {
                result.Append(' ');
            }
            // Full-width ASCII range (0xFF01-0xFF5E) to half-width (0x0021-0x007E)
            else if (ch >= '\uFF01' && ch <= '\uFF5E')
            {
                result.Append((char)(ch - 0xFEE0));
            }
            else
            {
                result.Append(ch);
            }
        }

        return CellValue.FromString(result.ToString());
    }
}
