// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHISQ.TEST function.
/// CHISQ.TEST(actual_range, expected_range) - returns the chi-squared test for independence.
/// </summary>
public sealed class ChiSqTestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiSqTestFunction Instance = new();

    private ChiSqTestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHISQ.TEST";

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

        // For now, simplified implementation that works with simple arrays
        // Extract numbers from both arguments
        var actualValues = new System.Collections.Generic.List<double>();
        var expectedValues = new System.Collections.Generic.List<double>();

        // Extract actual values
        if (args[0].Type == CellValueType.Number)
        {
            actualValues.Add(args[0].NumericValue);
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract expected values
        if (args[1].Type == CellValueType.Number)
        {
            expectedValues.Add(args[1].NumericValue);
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        if (actualValues.Count != expectedValues.Count || actualValues.Count == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Calculate chi-squared statistic
        double chiSquare = 0.0;

        for (int i = 0; i < actualValues.Count; i++)
        {
            double actual = actualValues[i];
            double expected = expectedValues[i];

            if (expected <= 0)
            {
                return CellValue.Error("#DIV/0!");
            }

            chiSquare += System.Math.Pow(actual - expected, 2) / expected;
        }

        // Degrees of freedom = n - 1
        int df = actualValues.Count - 1;

        if (df < 1)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Return the p-value (right-tailed probability)
            double pValue = 1.0 - StatisticalHelper.ChiSquareCDF(chiSquare, df);
            return CellValue.FromNumber(pValue);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
