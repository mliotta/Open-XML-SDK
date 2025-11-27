// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHOOSE function.
/// CHOOSE(index_num, value1, [value2], ...) - Returns value from list based on index.
/// </summary>
public sealed class ChooseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChooseFunction Instance = new();

    private ChooseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHOOSE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // CHOOSE requires at least 2 arguments (index and one value)
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // First argument must be the index
        var indexArg = args[0];
        if (indexArg.IsError)
        {
            return indexArg; // Propagate errors
        }

        if (indexArg.Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var index = (int)indexArg.NumericValue;

        // Index is 1-based in Excel
        // Index must be between 1 and the number of values
        if (index < 1 || index > args.Length - 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Return the value at the specified index
        // args[0] is the index, so the values start at args[1]
        return args[index];
    }
}
