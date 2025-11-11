// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COUNTBLANK function.
/// COUNTBLANK(range) - counts empty cells.
/// </summary>
public sealed class CountBlankFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CountBlankFunction Instance = new();

    private CountBlankFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COUNTBLANK";

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

            if (arg.Type == CellValueType.Empty)
            {
                count++;
            }
        }

        return CellValue.FromNumber(count);
    }
}
