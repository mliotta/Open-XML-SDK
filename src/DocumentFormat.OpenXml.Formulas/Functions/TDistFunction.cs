// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T.DIST function.
/// T.DIST(x, deg_freedom, cumulative) - returns the Student's t-distribution.
/// </summary>
public sealed class TDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TDistFunction Instance = new();

    private TDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T.DIST";

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
                result = StatisticalHelper.TDistCDF(x, df);
            }
            else
            {
                result = StatisticalHelper.TDistPDF(x, df);
            }

            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
