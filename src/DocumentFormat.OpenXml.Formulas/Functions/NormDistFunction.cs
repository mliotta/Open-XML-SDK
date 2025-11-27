// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORM.DIST function.
/// NORM.DIST(x, mean, standard_dev, cumulative) - returns the normal distribution.
/// </summary>
public sealed class NormDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormDistFunction Instance = new();

    private NormDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORM.DIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 4)
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

        // Get mean
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double mean = args[1].NumericValue;

        // Get standard deviation
        if (args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double standardDev = args[2].NumericValue;

        if (standardDev <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[3].Type == FormulaResultType.Boolean)
        {
            cumulative = args[3].BoolValue;
        }
        else if (args[3].Type == FormulaResultType.Number)
        {
            cumulative = args[3].NumericValue != 0;
        }
        else
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            double result = StatisticalHelper.NormDist(x, mean, standardDev, cumulative);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
