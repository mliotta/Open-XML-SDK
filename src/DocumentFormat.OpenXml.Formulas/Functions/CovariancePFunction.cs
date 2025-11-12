// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COVARIANCE.P function.
/// COVARIANCE.P(array1, array2) - Returns population covariance.
/// </summary>
public sealed class CovariancePFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CovariancePFunction Instance = new();

    private CovariancePFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COVARIANCE.P";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        var array1Values = new List<double>();
        var array2Values = new List<double>();

        // Extract numeric values from first array
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type == CellValueType.Number)
        {
            array1Values.Add(args[0].NumericValue);
        }

        // Extract numeric values from second array
        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[1].Type == CellValueType.Number)
        {
            array2Values.Add(args[1].NumericValue);
        }

        // Arrays must have same length
        if (array1Values.Count != array2Values.Count)
        {
            return CellValue.Error("#N/A");
        }

        // Need at least 1 data point
        if (array1Values.Count < 1)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate means
        var mean1 = array1Values.Average();
        var mean2 = array2Values.Average();

        // Calculate population covariance
        // Covariance.P = Σ((x-x̄)(y-ȳ)) / n
        var sumProduct = 0.0;

        for (int i = 0; i < array1Values.Count; i++)
        {
            var diff1 = array1Values[i] - mean1;
            var diff2 = array2Values[i] - mean2;
            sumProduct += diff1 * diff2;
        }

        var covariance = sumProduct / array1Values.Count;

        return CellValue.FromNumber(covariance);
    }
}
