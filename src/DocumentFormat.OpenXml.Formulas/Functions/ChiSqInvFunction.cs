// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHISQ.INV function.
/// CHISQ.INV(probability, deg_freedom) - returns the inverse of the left-tailed chi-squared distribution.
/// </summary>
public sealed class ChiSqInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChiSqInvFunction Instance = new();

    private ChiSqInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHISQ.INV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
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

        // Get degrees of freedom
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double df = args[1].NumericValue;

        if (df < 1 || df > 10000000000)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.ChiSquareInv(probability, df);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
