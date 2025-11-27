// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T.INV.2T function.
/// T.INV.2T(probability, deg_freedom) - returns the two-tailed inverse of the Student's t-distribution.
/// </summary>
public sealed class TInv2TFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TInv2TFunction Instance = new();

    private TInv2TFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T.INV.2T";

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

        // Get probability
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double probability = args[0].NumericValue;

        if (probability <= 0 || probability >= 1)
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
            // For two-tailed inverse, we want the value x such that P(|T| > x) = probability
            // This means P(T > x) = probability/2, so x = inverse(1 - probability/2)
            double result = StatisticalHelper.TDistInv(1.0 - probability / 2.0, df);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
