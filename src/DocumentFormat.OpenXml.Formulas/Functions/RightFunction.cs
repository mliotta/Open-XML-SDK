// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RIGHT function.
/// RIGHT(text, [num_chars]) - returns rightmost characters.
/// </summary>
public sealed class RightFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RightFunction Instance = new();

    private RightFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RIGHT";

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

        var startPos = System.Math.Max(0, text.Length - numChars);
        var result = text.Substring(startPos);
        return CellValue.FromString(result);
    }
}
