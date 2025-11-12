// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PHI function.
/// PHI(x) - returns the value of the density function for a standard normal distribution.
/// Formula: (1/√(2π)) * e^(-x²/2)
/// </summary>
public sealed class PhiFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PhiFunction Instance = new();

    private PhiFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PHI";

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

        // PHI(x) = (1/√(2π)) * e^(-x²/2)
        double result = StatisticalHelper.NormSPdf(x);
        return CellValue.FromNumber(result);
    }
}
