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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        double x = args[0].NumericValue;

        if (x <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.LogGamma(x);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
