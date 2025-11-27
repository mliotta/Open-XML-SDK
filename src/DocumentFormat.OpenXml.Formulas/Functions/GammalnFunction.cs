// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMALN function.
/// GAMMALN(x) - returns the natural logarithm of the gamma function.
/// </summary>
public sealed class GammalnFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammalnFunction Instance = new();

    private GammalnFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMALN";

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

        double x = args[0].NumericValue;

        if (x <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.LogGamma(x);
            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
