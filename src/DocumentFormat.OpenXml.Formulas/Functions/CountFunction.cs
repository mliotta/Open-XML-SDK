// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COUNT function.
/// COUNT(value1, [value2], ...) - counts cells with numbers.
/// </summary>
public sealed class CountFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CountFunction Instance = new();

    private CountFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COUNT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        var count = 0;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                count++;
            }
        }

        return CellValue.FromNumber(count);
    }
}
