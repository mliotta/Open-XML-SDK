// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COUNTA function.
/// COUNTA(value1, [value2], ...) - counts non-empty cells.
/// </summary>
public sealed class CountAFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CountAFunction Instance = new();

    private CountAFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COUNTA";

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

            if (arg.Type != CellValueType.Empty)
            {
                count++;
            }
        }

        return CellValue.FromNumber(count);
    }
}
