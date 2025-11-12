// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONFIDENCE function.
/// CONFIDENCE(alpha, standard_dev, size) - returns the confidence interval for a population mean (assumes normal distribution).
/// This is equivalent to CONFIDENCE.NORM.
/// </summary>
public sealed class ConfidenceFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConfidenceFunction Instance = new();

    private ConfidenceFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONFIDENCE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        // Get alpha (significance level)
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double alpha = args[0].NumericValue;

        if (alpha <= 0 || alpha >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Get standard deviation
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double standardDev = args[1].NumericValue;

        if (standardDev <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get sample size
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double size = args[2].NumericValue;

        if (size < 1)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Confidence interval = z * (stdev / sqrt(n))
            // where z is the critical value from standard normal distribution
            double z = StatisticalHelper.NormSInv(1 - alpha / 2);
            double result = z * (standardDev / System.Math.Sqrt(size));
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
