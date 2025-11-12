// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GEOMEAN function.
/// GEOMEAN(number1, [number2], ...) - returns the geometric mean of positive numbers.
/// </summary>
public sealed class GeomeanFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GeomeanFunction Instance = new();

    private GeomeanFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GEOMEAN";

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

        // Calculate geometric mean using logarithms to avoid overflow
        double logSum = 0.0;
        foreach (var value in values)
        {
            logSum += System.Math.Log(value);
        }

        double result = System.Math.Exp(logSum / values.Count);
        return CellValue.FromNumber(result);
    }
}
