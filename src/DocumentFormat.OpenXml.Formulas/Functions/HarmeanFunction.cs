// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HARMEAN function.
/// HARMEAN(number1, [number2], ...) - returns the harmonic mean of positive numbers.
/// </summary>
public sealed class HarmeanFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HarmeanFunction Instance = new();

    private HarmeanFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HARMEAN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var values = new List<double>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }

            if (arg.Type == CellValueType.Number)
            {
                if (arg.NumericValue <= 0)
                {
                    return CellValue.Error("#NUM!");
                }
                values.Add(arg.NumericValue);
            }
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Calculate harmonic mean: n / sum(1/xi)
        double reciprocalSum = 0.0;
        foreach (var value in values)
        {
            reciprocalSum += 1.0 / value;
        }

        double result = values.Count / reciprocalSum;
        return CellValue.FromNumber(result);
    }
}
