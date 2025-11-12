// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SERIESSUM function.
/// SERIESSUM(x, n, m, coefficients) - returns the sum of a power series.
/// Formula: Σ(coefficient[i] * x^(n + m*i)) for i = 0 to k-1
/// </summary>
public sealed class SeriesSumFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SeriesSumFunction Instance = new();

    private SeriesSumFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SERIESSUM";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in all arguments
        for (int i = 0; i < 4; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Validate first three arguments are numbers
        if (args[0].Type != CellValueType.Number ||
            args[1].Type != CellValueType.Number ||
            args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var x = args[0].NumericValue;
        var n = args[1].NumericValue;
        var m = args[2].NumericValue;

        // Fourth argument should be a number (representing a single coefficient)
        // In a full implementation, this would handle arrays
        if (args[3].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        // For single coefficient, the formula is: coefficient * x^n
        var coefficient = args[3].NumericValue;

        // Calculate x^n
        double power;
        if (x == 0 && n < 0)
        {
            return CellValue.Error("#NUM!"); // Division by zero
        }

        try
        {
            power = System.Math.Pow(x, n);
        }
        catch
        {
            return CellValue.Error("#NUM!");
        }

        var result = coefficient * power;

        // Check for overflow or invalid result
        if (double.IsInfinity(result) || double.IsNaN(result))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(result);
    }
}
