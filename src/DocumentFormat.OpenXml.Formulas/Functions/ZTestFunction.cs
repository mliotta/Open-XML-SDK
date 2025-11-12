// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the Z.TEST function.
/// Z.TEST(array, x, [sigma]) - returns the one-tailed P-value of a z-test.
/// </summary>
public sealed class ZTestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ZTestFunction Instance = new();

    private ZTestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "Z.TEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Get x value (the hypothesized population mean)
        double x;
        double? sigma = null;

        // Determine parameter positions
        int xIndex = args.Length >= 2 ? args.Length - 2 : -1;
        int sigmaIndex = args.Length >= 3 ? args.Length - 1 : -1;

        // Check if last argument could be sigma
        if (sigmaIndex >= 0 && args[sigmaIndex].Type == CellValueType.Number)
        {
            sigma = args[sigmaIndex].NumericValue;
            if (sigma <= 0)
            {
                return CellValue.Error("#NUM!");
            }
        }
        else
        {
            sigmaIndex = -1;
            xIndex = args.Length - 1;
        }

        if (args[xIndex].IsError)
        {
            return args[xIndex];
        }

        if (args[xIndex].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        x = args[xIndex].NumericValue;

        // Collect array values
        var values = new List<double>();
        int endIndex = sigmaIndex >= 0 ? sigmaIndex : xIndex;

        for (int i = 0; i < endIndex; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }

            if (args[i].Type == CellValueType.Number)
            {
                values.Add(args[i].NumericValue);
            }
        }

        if (values.Count == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Calculate sample mean
        double sum = 0;
        foreach (var value in values)
        {
            sum += value;
        }
        double mean = sum / values.Count;

        // Calculate or use provided standard deviation
        double stdDev;
        if (sigma.HasValue)
        {
            stdDev = sigma.Value;
        }
        else
        {
            // Calculate sample standard deviation
            double sumSquaredDiff = 0;
            foreach (var value in values)
            {
                double diff = value - mean;
                sumSquaredDiff += diff * diff;
            }
            stdDev = System.Math.Sqrt(sumSquaredDiff / values.Count);
        }

        if (stdDev == 0)
        {
            return CellValue.Error("#DIV/0!");
        }

        // Calculate z-score
        double z = (mean - x) / (stdDev / System.Math.Sqrt(values.Count));

        // Return one-tailed P-value (upper tail)
        double pValue = 1.0 - StatisticalHelper.NormSDist(z);

        return CellValue.FromNumber(pValue);
    }
}
