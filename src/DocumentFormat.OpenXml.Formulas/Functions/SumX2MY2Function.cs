// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUMX2MY2 function.
/// SUMX2MY2(array_x, array_y) - returns the sum of differences of squares (Σ(x²-y²)).
/// </summary>
public sealed class SumX2MY2Function : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumX2MY2Function Instance = new();

    private SumX2MY2Function()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUMX2MY2";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
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
        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        // For single values, compute x² - y²
        var x = args[0].NumericValue;
        var y = args[1].NumericValue;
        var result = (x * x) - (y * y);

        // Check for overflow
        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
