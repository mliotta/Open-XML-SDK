// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NEGBINOM.DIST function.
/// NEGBINOM.DIST(number_f, number_s, probability_s, cumulative) - returns the negative binomial distribution.
/// </summary>
public sealed class NegBinomDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NegBinomDistFunction Instance = new();

    private NegBinomDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NEGBINOM.DIST";

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

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number ||
            args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        int numberF = (int)args[0].NumericValue;
        int numberS = (int)args[1].NumericValue;
        double probabilityS = args[2].NumericValue;

        if (numberF < 0 || numberS < 1 || probabilityS < 0 || probabilityS > 1)
        {
            return CellValue.Error("#NUM!");
        }

        bool cumulative;
        if (args[3].Type == CellValueType.Boolean)
        {
            cumulative = args[3].BoolValue;
        }
        else if (args[3].Type == CellValueType.Number)
        {
            cumulative = args[3].NumericValue != 0;
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
                // CDF: Sum of PMF from 0 to numberF
                result = 0;
                for (int k = 0; k <= numberF; k++)
                {
                    result += NegBinomPMF(k, numberS, probabilityS);
                }
            }
            else
            {
                // PMF: C(k+r-1, k) * p^r * (1-p)^k
                result = NegBinomPMF(numberF, numberS, probabilityS);
            }

            return CellValue.FromNumber(result);
        }
        catch (System.Exception)
        {
            return CellValue.Error("#NUM!");
        }
    }

    private double NegBinomPMF(int k, int r, double p)
    {
        if (p == 1.0)
            return k == 0 ? 1.0 : 0.0;

        // C(k+r-1, k) * p^r * (1-p)^k
        double logBinom = StatisticalHelper.LogGamma(k + r) -
                         StatisticalHelper.LogGamma(k + 1) -
                         StatisticalHelper.LogGamma(r);
        double logProb = r * System.Math.Log(p) + k * System.Math.Log(1.0 - p);
        return System.Math.Exp(logBinom + logProb);
    }
}
