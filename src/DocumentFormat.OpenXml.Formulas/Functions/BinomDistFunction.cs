// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BINOM.DIST function.
/// BINOM.DIST(number_s, trials, probability_s, cumulative) - returns the binomial distribution probability.
/// </summary>
public sealed class BinomDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BinomDistFunction Instance = new();

    private BinomDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BINOM.DIST";

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

        int numberS = (int)args[0].NumericValue;
        int trials = (int)args[1].NumericValue;
        double probabilityS = args[2].NumericValue;

        if (numberS < 0 || trials < 0 || numberS > trials || probabilityS < 0 || probabilityS > 1)
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
                result = StatisticalHelper.BinomialCDF(numberS, trials, probabilityS);
            }
            else
            {
                result = StatisticalHelper.BinomialPMF(numberS, trials, probabilityS);
            }

            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
