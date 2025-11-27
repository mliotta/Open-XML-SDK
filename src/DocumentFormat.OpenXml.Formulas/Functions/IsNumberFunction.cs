// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISNUMBER function.
/// ISNUMBER(value) - TRUE if value is a number.
/// </summary>
public sealed class IsNumberFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsNumberFunction Instance = new();

    private IsNumberFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISNUMBER";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        var isNumber = args[0].Type == FormulaResultType.Number;
        return FormulaResult.FromBool(isNumber);
    }
}
