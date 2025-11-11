// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CODE function.
/// CODE(text) - returns ASCII code for first character.
/// </summary>
public sealed class CodeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CodeFunction Instance = new();

    private CodeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CODE";

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

        // Get the code of the first character
        var code = (int)text[0];

        return CellValue.FromNumber(code);
    }
}
