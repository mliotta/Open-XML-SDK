// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORM.INV function.
/// NORM.INV(probability, mean, standard_dev) - returns the inverse of the normal cumulative distribution.
/// </summary>
public sealed class NormInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormInvFunction Instance = new();

    private NormInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORM.INV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        if (probability <= 0 || probability >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Get mean
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double mean = args[1].NumericValue;

        // Get standard deviation
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double standardDev = args[2].NumericValue;

        if (standardDev <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.NormInv(probability, mean, standardDev);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
