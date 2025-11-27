// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHISQ.DIST function.
/// CHISQ.DIST(x, deg_freedom, cumulative) - returns the chi-squared distribution.
/// </summary>
public sealed class ChiSqDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiSqDistFunction Instance = new();

    private ChiSqDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHISQ.DIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
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

        // Get cumulative flag
        bool cumulative;
        if (args[2].Type == FormulaResultType.Boolean)
        {
            cumulative = args[2].BoolValue;
        }
        else if (args[2].Type == FormulaResultType.Number)
        {
            cumulative = args[2].NumericValue != 0;
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
                result = StatisticalHelper.ChiSquareCDF(x, df);
            }
            else
            {
                result = StatisticalHelper.ChiSquarePDF(x, df);
            }

            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
