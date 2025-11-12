// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REVERSE function.
/// REVERSE(text) - reverses text string (bonus function, not standard Excel).
/// </summary>
public sealed class ReverseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ReverseFunction Instance = new();

    private ReverseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "REVERSE";

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
        var charArray = text.ToCharArray();
        Array.Reverse(charArray);
        var result = new string(charArray);

        return CellValue.FromString(result);
    }
}
