// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRIMALL function.
/// TRIMALL(text) - removes all spaces from text.
/// </summary>
public sealed class TrimAllFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrimAllFunction Instance = new();

    private TrimAllFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRIMALL";

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
            return CellValue.FromString(string.Empty);
        }

        // Remove all spaces from text
        var sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c != ' ')
            {
                sb.Append(c);
            }
        }

        return CellValue.FromString(sb.ToString());
    }
}
