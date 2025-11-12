// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the F.TEST function.
/// F.TEST(array1, array2) - returns the result of an F-test.
/// </summary>
public sealed class FTestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FTestFunction Instance = new();

    private FTestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "F.TEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        // Extract numeric values from arrays
        var values1 = ExtractNumericValues(args[0]);
        var values2 = ExtractNumericValues(args[1]);

        if (values1.Count < 2 || values2.Count < 2)
        {
            return CellValue.Error("#DIV/0!");
        }

        try
        {
            // Calculate variances
            double mean1 = values1.Average();
            double variance1 = values1.Sum(v => System.Math.Pow(v - mean1, 2)) / (values1.Count - 1);

            double mean2 = values2.Average();
            double variance2 = values2.Sum(v => System.Math.Pow(v - mean2, 2)) / (values2.Count - 1);

            if (variance1 <= 0 || variance2 <= 0)
            {
                return CellValue.Error("#DIV/0!");
            }

            // Calculate F statistic
            double f = variance1 / variance2;
            int df1 = values1.Count - 1;
            int df2 = values2.Count - 1;

            // Calculate two-tailed p-value
            double cdf = StatisticalHelper.FDistCDF(f, df1, df2);
            double pValue = 2.0 * System.Math.Min(cdf, 1.0 - cdf);

            return CellValue.FromNumber(pValue);
        }
        catch (System.Exception)
        {
            return CellValue.Error("#NUM!");
        }
    }

    private List<double> ExtractNumericValues(CellValue arg)
    {
        var values = new List<double>();

        if (arg.Type == CellValueType.Number)
        {
            values.Add(arg.NumericValue);
        }

        return values;
    }
}
