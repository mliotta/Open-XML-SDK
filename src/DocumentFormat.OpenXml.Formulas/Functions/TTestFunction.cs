// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T.TEST function.
/// T.TEST(array1, array2, tails, type) - returns the probability associated with a Student's t-test.
/// </summary>
public sealed class TTestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TTestFunction Instance = new();

    private TTestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T.TEST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        var array1Values = new List<double>();
        var array2Values = new List<double>();

        // Extract numeric values from arrays
        if (args[0].Type == CellValueType.Number)
        {
            array1Values.Add(args[0].NumericValue);
        }

        if (args[1].Type == CellValueType.Number)
        {
            array2Values.Add(args[1].NumericValue);
        }

        if (array1Values.Count == 0 || array2Values.Count == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Get tails parameter
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        int tails = (int)args[2].NumericValue;

        if (tails != 1 && tails != 2)
        {
            return CellValue.Error("#NUM!");
        }

        // Get type parameter
        if (args[3].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        int type = (int)args[3].NumericValue;

        if (type < 1 || type > 3)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double tStat;
            double df;

            if (type == 1)
            {
                // Paired t-test
                if (array1Values.Count != array2Values.Count)
                {
                    return CellValue.Error("#N/A");
                }

                var differences = new List<double>();
                for (int i = 0; i < array1Values.Count; i++)
                {
                    differences.Add(array1Values[i] - array2Values[i]);
                }

                var meanDiff = differences.Average();
                var n = differences.Count;
                var variance = differences.Sum(d => System.Math.Pow(d - meanDiff, 2)) / (n - 1);
                var se = System.Math.Sqrt(variance / n);

                if (se == 0)
                {
                    return CellValue.Error("#DIV/0!");
                }

                tStat = System.Math.Abs(meanDiff / se);
                df = n - 1;
            }
            else if (type == 2)
            {
                // Two-sample equal variance
                var n1 = array1Values.Count;
                var n2 = array2Values.Count;
                var mean1 = array1Values.Average();
                var mean2 = array2Values.Average();

                var var1 = array1Values.Sum(v => System.Math.Pow(v - mean1, 2));
                var var2 = array2Values.Sum(v => System.Math.Pow(v - mean2, 2));

                var pooledVariance = (var1 + var2) / (n1 + n2 - 2);
                var se = System.Math.Sqrt(pooledVariance * (1.0 / n1 + 1.0 / n2));

                if (se == 0)
                {
                    return CellValue.Error("#DIV/0!");
                }

                tStat = System.Math.Abs((mean1 - mean2) / se);
                df = n1 + n2 - 2;
            }
            else
            {
                // Two-sample unequal variance (Welch's t-test)
                var n1 = array1Values.Count;
                var n2 = array2Values.Count;
                var mean1 = array1Values.Average();
                var mean2 = array2Values.Average();

                var var1 = array1Values.Sum(v => System.Math.Pow(v - mean1, 2)) / (n1 - 1);
                var var2 = array2Values.Sum(v => System.Math.Pow(v - mean2, 2)) / (n2 - 1);

                var se = System.Math.Sqrt(var1 / n1 + var2 / n2);

                if (se == 0)
                {
                    return CellValue.Error("#DIV/0!");
                }

                tStat = System.Math.Abs((mean1 - mean2) / se);

                // Welch-Satterthwaite degrees of freedom
                var numerator = System.Math.Pow(var1 / n1 + var2 / n2, 2);
                var denominator = System.Math.Pow(var1 / n1, 2) / (n1 - 1) + System.Math.Pow(var2 / n2, 2) / (n2 - 1);
                df = numerator / denominator;
            }

            if (df <= 0)
            {
                return CellValue.Error("#NUM!");
            }

            // Calculate p-value
            double pValue;
            if (tails == 1)
            {
                // One-tailed test
                pValue = 1.0 - StatisticalHelper.TDistCDF(tStat, df);
            }
            else
            {
                // Two-tailed test
                pValue = 2.0 * (1.0 - StatisticalHelper.TDistCDF(tStat, df));
            }

            return CellValue.FromNumber(pValue);
        }
        catch (System.Exception)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
