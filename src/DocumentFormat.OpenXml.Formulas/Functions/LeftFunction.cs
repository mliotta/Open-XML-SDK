// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LEFT function.
/// LEFT(text, [num_chars]) - returns leftmost characters.
/// </summary>
public sealed class LeftFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LeftFunction Instance = new();

    private LeftFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LEFT";

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
        var numChars = 1;

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

            numChars = (int)args[1].NumericValue;

            if (numChars < 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var result = numChars >= text.Length ? text : text.Substring(0, numChars);
        return CellValue.FromString(result);
    }
}
