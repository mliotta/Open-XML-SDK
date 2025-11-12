// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUMSQ function.
/// SUMSQ(number1, [number2], ...) - returns the sum of the squares of the arguments.
/// </summary>
public sealed class SumSqFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumSqFunction Instance = new();

    private SumSqFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUMSQ";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var sumOfSquares = 0.0;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                var value = arg.NumericValue;
                sumOfSquares += value * value;

                // Check for overflow
                if (double.IsInfinity(sumOfSquares) || double.IsNaN(sumOfSquares))
                {
                    return CellValue.Error("#NUM!");
                }
            }
        }

        return CellValue.FromNumber(sumOfSquares);
    }
}
