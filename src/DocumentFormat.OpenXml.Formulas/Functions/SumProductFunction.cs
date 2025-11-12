// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SUMPRODUCT function.
/// SUMPRODUCT(array1, [array2], ...) - returns the sum of products of corresponding array elements.
/// </summary>
public sealed class SumProductFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SumProductFunction Instance = new();

    private SumProductFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SUMPRODUCT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Collect all numeric values from each array
        var arrays = new System.Collections.Generic.List<double[]>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                arrays.Add(new[] { arg.NumericValue });
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Check if all arrays are non-empty
        if (arrays.Count == 0)
        {
            return CellValue.FromNumber(0);
        }

        // Get the length (all arrays should have same length in Excel, but we'll handle single values)
        var length = arrays[0].Length;
        foreach (var array in arrays)
        {
            if (array.Length != length && array.Length != 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Calculate sum of products
        var sum = 0.0;
        for (int i = 0; i < length; i++)
        {
            var product = 1.0;
            foreach (var array in arrays)
            {
                var index = array.Length == 1 ? 0 : i;
                product *= array[index];
            }

            sum += product;
        }

        return CellValue.FromNumber(sum);
    }
}
