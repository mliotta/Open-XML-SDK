// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HSTACK function.
/// HSTACK(array1, [array2], ...) - Stacks arrays horizontally (column-wise).
/// NOTE: Due to single-value return limitation, only the first element of the stacked array is returned.
/// </summary>
public sealed class HStackFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HStackFunction Instance = new();

    private HStackFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HSTACK";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in any argument
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        // Return the first element (top-left of the stacked result)
        // In a full implementation, this would stack all arrays horizontally
        // and return the entire result array
        return args[0];
    }
}
