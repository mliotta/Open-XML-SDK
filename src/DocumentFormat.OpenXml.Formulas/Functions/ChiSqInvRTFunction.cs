// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHISQ.INV.RT function.
/// CHISQ.INV.RT(probability, deg_freedom) - returns the inverse of the right-tailed chi-squared distribution.
/// </summary>
public sealed class ChiSqInvRTFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiSqInvRTFunction Instance = new();

    private ChiSqInvRTFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHISQ.INV.RT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
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

        // Get probability
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double probability = args[0].NumericValue;

        if (probability < 0 || probability > 1)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get degrees of freedom
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double df = args[1].NumericValue;

        if (df < 1 || df > 10000000000)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Right-tailed inverse: find x such that P(X > x) = probability
            // This is equivalent to P(X < x) = 1 - probability
            double result = StatisticalHelper.ChiSquareInv(1.0 - probability, df);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
