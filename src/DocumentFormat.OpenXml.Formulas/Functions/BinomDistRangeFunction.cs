// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BINOM.DIST.RANGE function.
/// BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2]) - Binomial probability for range of trials.
/// </summary>
public sealed class BinomDistRangeFunction : IFunctionImplementation
{
    public static readonly BinomDistRangeFunction Instance = new();

    private BinomDistRangeFunction()
    {
    }

    public string Name => "BINOM.DIST.RANGE";

    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 4)
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

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var trials = (int)args[0].NumericValue;
        var prob = args[1].NumericValue;
        var numberS = (int)args[2].NumericValue;
        var numberS2 = args.Length == 4 && args[3].Type == CellValueType.Number ? (int)args[3].NumericValue : numberS;

        if (trials < 0 || prob < 0 || prob > 1 || numberS < 0 || numberS2 < 0 || numberS > trials || numberS2 > trials)
        {
            return CellValue.Error("#NUM!");
        }

        if (numberS > numberS2)
        {
            return CellValue.Error("#NUM!");
        }

        var sum = 0.0;
        for (var k = numberS; k <= numberS2; k++)
        {
            sum += StatisticalHelper.BinomialPMF(trials, k, prob);
        }

        return CellValue.FromNumber(sum);
    }
}
