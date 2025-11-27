// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUMXMY2 function.
/// SUMXMY2(array_x, array_y) - returns the sum of squares of differences (Σ(x-y)²).
/// </summary>
public sealed class SumXMY2Function : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumXMY2Function Instance = new();

    private SumXMY2Function()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUMXMY2";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        // Both arguments must be numbers for single-value case
        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // For single values, compute (x - y)²
        var x = args[0].NumericValue;
        var y = args[1].NumericValue;
        var diff = x - y;
        var result = diff * diff;

        // Check for overflow
        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(result);
    }
}
