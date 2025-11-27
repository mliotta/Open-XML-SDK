// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TDIST function (legacy compatibility).
/// TDIST(x, deg_freedom, tails) - returns the Student's t-distribution (Excel 2007 compatibility).
/// </summary>
public sealed class TDistLegacyFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TDistLegacyFunction Instance = new();

    private TDistLegacyFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TDIST";

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

        // TDIST requires x >= 0
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

        // Get tails (1 or 2)
        if (args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        int tails = (int)args[2].NumericValue;

        if (tails != 1 && tails != 2)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            // TDIST always returns right-tailed probability
            double rightTail = 1.0 - StatisticalHelper.TDistCDF(x, df);

            double result = tails == 1 ? rightTail : 2.0 * rightTail;
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
