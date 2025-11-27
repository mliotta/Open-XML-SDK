// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHISQ.DIST.RT function.
/// CHISQ.DIST.RT(x, deg_freedom) - returns the right-tailed chi-squared distribution.
/// </summary>
public sealed class ChiSqDistRTFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiSqDistRTFunction Instance = new();

    private ChiSqDistRTFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHISQ.DIST.RT";

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

        // Get x value
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double x = args[0].NumericValue;

        if (x < 0)
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
            // Right-tailed: P(X > x) = 1 - CDF(x)
            double result = 1.0 - StatisticalHelper.ChiSquareCDF(x, df);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
