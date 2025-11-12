// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BETA.DIST function.
/// BETA.DIST(x, alpha, beta, cumulative, [A], [B]) - returns the beta cumulative distribution function.
/// </summary>
public sealed class BetaDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BetaDistFunction Instance = new();

    private BetaDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BETA.DIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 4 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
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
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double x = args[0].NumericValue;

        // Get alpha
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double alpha = args[1].NumericValue;

        if (alpha <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get beta
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double beta = args[2].NumericValue;

        if (beta <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get cumulative flag
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

        // Get optional A (lower bound)
        double A = 0.0;
        if (args.Length > 4 && args[4].Type == CellValueType.Number)
        {
            A = args[4].NumericValue;
        }

        // Get optional B (upper bound)
        double B = 1.0;
        if (args.Length > 5 && args[5].Type == CellValueType.Number)
        {
            B = args[5].NumericValue;
        }

        if (A >= B)
        {
            return CellValue.Error("#NUM!");
        }

        if (x < A || x > B)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Transform x to [0, 1] range
            double xTransformed = (x - A) / (B - A);

            double result;
            if (cumulative)
            {
                result = StatisticalHelper.BetaCDF(xTransformed, alpha, beta);
            }
            else
            {
                // PDF needs to be scaled by 1/(B-A)
                result = StatisticalHelper.BetaPDF(xTransformed, alpha, beta) / (B - A);
            }

            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
