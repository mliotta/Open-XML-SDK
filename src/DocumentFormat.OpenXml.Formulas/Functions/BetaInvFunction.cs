// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the BETA.INV function.
/// BETA.INV(probability, alpha, beta, [A], [B]) - returns the inverse of the beta cumulative distribution function.
/// </summary>
public sealed class BetaInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly BetaInvFunction Instance = new();

    private BetaInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "BETA.INV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 5)
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

        // Get probability
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double probability = args[0].NumericValue;

        if (probability < 0 || probability > 1)
        {
            return CellValue.Error("#NUM!");
        }

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

        // Get optional A (lower bound)
        double A = 0.0;
        if (args.Length > 3 && args[3].Type == CellValueType.Number)
        {
            A = args[3].NumericValue;
        }

        // Get optional B (upper bound)
        double B = 1.0;
        if (args.Length > 4 && args[4].Type == CellValueType.Number)
        {
            B = args[4].NumericValue;
        }

        if (A >= B)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Get inverse in [0, 1] range
            double xTransformed = StatisticalHelper.BetaInv(probability, alpha, beta);

            // Transform back to [A, B] range
            double result = A + xTransformed * (B - A);

            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
