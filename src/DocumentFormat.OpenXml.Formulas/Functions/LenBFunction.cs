// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LENB function.
/// LENB(text) - returns the length of text in bytes (UTF-8).
/// </summary>
public sealed class LenBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LenBFunction Instance = new();

    private LenBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LENB";

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
        var byteCount = Encoding.UTF8.GetByteCount(text);

        return CellValue.FromNumber(byteCount);
    }
}
