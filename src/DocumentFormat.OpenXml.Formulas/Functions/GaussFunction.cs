// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAUSS function.
/// GAUSS(z) - returns the probability that a member of a standard normal population will fall between the mean and z standard deviations from the mean.
/// Formula: NORM.S.DIST(z, TRUE) - 0.5
/// </summary>
public sealed class GaussFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GaussFunction Instance = new();

    private GaussFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAUSS";

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

        double z = args[0].NumericValue;

        // GAUSS(z) = NORM.S.DIST(z, TRUE) - 0.5
        double result = StatisticalHelper.NormSDist(z) - 0.5;
        return CellValue.FromNumber(result);
    }
}
