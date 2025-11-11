// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CLEAN function.
/// CLEAN(text) - removes non-printable characters.
/// </summary>
public sealed class CleanFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CleanFunction Instance = new();

    private CleanFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CLEAN";

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

        // Remove non-printable characters (ASCII codes 0-31)
        var sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c >= 32)
            {
                sb.Append(c);
            }
        }

        return CellValue.FromString(sb.ToString());
    }
}
