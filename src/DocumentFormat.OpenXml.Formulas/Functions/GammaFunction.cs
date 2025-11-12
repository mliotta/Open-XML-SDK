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

        double number = args[0].NumericValue;

        // GAMMA function is not defined for zero and negative integers
        if (number <= 0 && number == System.Math.Floor(number))
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.Gamma(number);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
