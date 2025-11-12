// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DBCS function.
/// DBCS(text) - Converts half-width (single-byte) characters to full-width (double-byte) characters.
/// Used primarily for Japanese text.
/// </summary>
public sealed class DbcsFunction : IFunctionImplementation
{
    public static readonly DbcsFunction Instance = new();

    private DbcsFunction()
    {
    }

    public string Name => "DBCS";

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
            // Half-width space to full-width space
            if (ch == ' ')
            {
                result.Append('\u3000');
            }
            // Half-width ASCII range (0x0021-0x007E) to full-width (0xFF01-0xFF5E)
            else if (ch >= '!' && ch <= '~')
            {
                result.Append((char)(ch + 0xFEE0));
            }
            else
            {
                result.Append(ch);
            }
        }

        return CellValue.FromString(result.ToString());
    }
}
