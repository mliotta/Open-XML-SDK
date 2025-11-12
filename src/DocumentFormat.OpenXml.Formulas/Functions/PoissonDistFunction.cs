// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the POISSON.DIST function.
/// POISSON.DIST(x, mean, cumulative) - returns the Poisson distribution.
/// </summary>
public sealed class PoissonDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PoissonDistFunction Instance = new();

    private PoissonDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "POISSON.DIST";

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

        // Get x value
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        int x = (int)System.Math.Floor(args[0].NumericValue);

        if (x < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get mean
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double mean = args[1].NumericValue;

        if (mean <= 0.0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[2].Type == CellValueType.Boolean)
        {
            cumulative = args[2].BoolValue;
        }
        else if (args[2].Type == CellValueType.Number)
        {
            cumulative = args[2].NumericValue != 0;
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            double result;
            if (cumulative)
            {
                // CDF: Sum of probabilities from 0 to x
                result = 0.0;
                for (int i = 0; i <= x; i++)
                {
                    result += PoissonPMF(i, mean);
                }
            }
            else
            {
                // PMF: (mean^x * exp(-mean)) / x!
                result = PoissonPMF(x, mean);
            }

            return CellValue.FromNumber(result);
        }
        catch (System.Exception)
        {
            return CellValue.Error("#NUM!");
        }
    }

    private double PoissonPMF(int x, double mean)
    {
        // Use logarithms to avoid overflow
        double logProb = x * System.Math.Log(mean) - mean - StatisticalHelper.LogGamma(x + 1.0);
        return System.Math.Exp(logProb);
    }
}
