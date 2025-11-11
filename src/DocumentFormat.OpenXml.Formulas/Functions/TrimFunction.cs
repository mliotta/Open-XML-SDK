// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRIM function.
/// TRIM(text) - removes extra spaces.
/// </summary>
public sealed class TrimFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrimFunction Instance = new();

    private TrimFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRIM";

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

        // Remove leading/trailing spaces and collapse internal spaces
        text = text.Trim();
        text = Regex.Replace(text, @"\s+", " ");

        return CellValue.FromString(text);
    }
}
