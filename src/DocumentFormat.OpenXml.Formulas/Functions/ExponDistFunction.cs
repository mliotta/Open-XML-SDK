// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EXPON.DIST function.
/// EXPON.DIST(x, lambda, cumulative) - returns the exponential distribution.
/// </summary>
public sealed class ExponDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ExponDistFunction Instance = new();

    private ExponDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EXPON.DIST";

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

        // Get lambda (rate parameter)
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double lambda = args[1].NumericValue;

        if (lambda <= 0)
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

        double result;
        if (cumulative)
        {
            // CDF: 1 - exp(-lambda * x)
            result = 1.0 - System.Math.Exp(-lambda * x);
        }
        else
        {
            // PDF: lambda * exp(-lambda * x)
            result = lambda * System.Math.Exp(-lambda * x);
        }

        return FormulaResult.FromNumber(result);
    }
}
