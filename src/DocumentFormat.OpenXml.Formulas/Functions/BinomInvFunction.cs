// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BINOM.INV function.
/// BINOM.INV(trials, probability_s, alpha) - returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
/// </summary>
public sealed class BinomInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BinomInvFunction Instance = new();

    private BinomInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BINOM.INV";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
        {
            return FormulaResult.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number ||
            args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        int trials = (int)args[0].NumericValue;
        double probabilityS = args[1].NumericValue;
        double alpha = args[2].NumericValue;

        if (trials < 0 || probabilityS < 0 || probabilityS > 1 || alpha < 0 || alpha > 1)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Find the smallest value k where CDF(k) >= alpha
            for (int k = 0; k <= trials; k++)
            {
                double cdf = StatisticalHelper.BinomialCDF(k, trials, probabilityS);
                if (cdf >= alpha)
                {
                    return FormulaResult.FromNumber(k);
                }
            }

            // Should not reach here if alpha <= 1, but return trials as fallback
            return FormulaResult.FromNumber(trials);
        }
        catch (System.Exception)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
