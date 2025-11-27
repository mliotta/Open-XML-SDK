// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISODD function.
/// ISODD(number) - TRUE if number is odd.
/// </summary>
public sealed class IsOddFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsOddFunction Instance = new();

    private IsOddFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISODD";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Note: Errors are NOT propagated for IS* functions
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.FromBool(false);
        }

        var number = args[0].NumericValue;
        var truncated = System.Math.Truncate(number);
        var isOdd = truncated % 2 != 0;
        return FormulaResult.FromBool(isOdd);
    }
}
