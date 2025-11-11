// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FIND function.
/// FIND(find_text, within_text, [start_num]) - finds text (case-sensitive, 1-based).
/// </summary>
public sealed class FindFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FindFunction Instance = new();

    private FindFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FIND";

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

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= withinText.Length)
        {
            return CellValue.Error("#VALUE!");
        }

        var position = withinText.IndexOf(findText, startIndex, StringComparison.Ordinal);

        if (position == -1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Return 1-based position
        return CellValue.FromNumber(position + 1);
    }
}
