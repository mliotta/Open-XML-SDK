// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMA function.
/// GAMMA(number) - returns the gamma function value.
/// </summary>
public sealed class GammaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammaFunction Instance = new();

    private GammaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMA";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        double number = args[0].NumericValue;

        // GAMMA function is not defined for zero and negative integers
        if (number <= 0 && number == System.Math.Floor(number))
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.Gamma(number);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
