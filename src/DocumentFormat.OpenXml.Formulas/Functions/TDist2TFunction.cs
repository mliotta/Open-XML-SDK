// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T.DIST.2T function.
/// T.DIST.2T(x, deg_freedom) - returns the two-tailed Student's t-distribution.
/// </summary>
public sealed class TDist2TFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TDist2TFunction Instance = new();

    private TDist2TFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T.DIST.2T";

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

        // For two-tailed test, x must be non-negative
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

        if (df < 1)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // Two-tailed: P(|T| > x) = 2 * P(T > x) = 2 * (1 - CDF(x))
            double result = 2.0 * (1.0 - StatisticalHelper.TDistCDF(x, df));
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
