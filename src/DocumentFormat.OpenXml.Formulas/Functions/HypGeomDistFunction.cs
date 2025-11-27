// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HYPGEOM.DIST function.
/// HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative) - returns the hypergeometric distribution.
/// </summary>
public sealed class HypGeomDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HypGeomDistFunction Instance = new();

    private HypGeomDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HYPGEOM.DIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 5)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in arguments
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        // Get sample_s (number of successes in sample)
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        int x = (int)System.Math.Floor(args[0].NumericValue);

        // Get number_sample (size of sample)
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        int n = (int)System.Math.Floor(args[1].NumericValue);

        // Get population_s (number of successes in population)
        if (args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        int K = (int)System.Math.Floor(args[2].NumericValue);

        // Get number_pop (population size)
        if (args[3].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        int N = (int)System.Math.Floor(args[3].NumericValue);

        // Validate parameters
        if (x < 0 || n < 0 || K < 0 || N < 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        if (x > n || x > K || n > N || K > N)
        {
            return FormulaResult.Error("#NUM!");
        }

        if (x < System.Math.Max(0, n + K - N))
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[4].Type == FormulaResultType.Boolean)
        {
            cumulative = args[4].BoolValue;
        }
        else if (args[4].Type == FormulaResultType.Number)
        {
            cumulative = args[4].NumericValue != 0;
        }
        else
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            double result;
            if (cumulative)
            {
                // CDF: Sum of probabilities from max(0, n+K-N) to x
                result = 0.0;
                int minX = System.Math.Max(0, n + K - N);
                for (int i = minX; i <= x; i++)
                {
                    result += HypergeometricPMF(i, n, K, N);
                }
            }
            else
            {
                // PMF: C(K,x) * C(N-K, n-x) / C(N, n)
                result = HypergeometricPMF(x, n, K, N);
            }

            return FormulaResult.FromNumber(result);
        }
        catch (System.Exception)
        {
            return FormulaResult.Error("#NUM!");
        }
    }

    private double HypergeometricPMF(int x, int n, int K, int N)
    {
        // Use logarithms to avoid overflow
        // PMF = C(K,x) * C(N-K, n-x) / C(N, n)
        double logNumer = LogCombination(K, x) + LogCombination(N - K, n - x);
        double logDenom = LogCombination(N, n);
        return System.Math.Exp(logNumer - logDenom);
    }

    private double LogCombination(int n, int k)
    {
        if (k < 0 || k > n)
            return double.NegativeInfinity;
        if (k == 0 || k == n)
            return 0.0;

        return StatisticalHelper.LogGamma(n + 1.0) -
               StatisticalHelper.LogGamma(k + 1.0) -
               StatisticalHelper.LogGamma(n - k + 1.0);
    }
}
