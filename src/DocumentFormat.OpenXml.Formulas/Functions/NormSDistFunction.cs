// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORM.S.DIST function.
/// NORM.S.DIST(z, cumulative) - returns the standard normal distribution (mean=0, stdev=1).
/// </summary>
public sealed class NormSDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormSDistFunction Instance = new();

    private NormSDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORM.S.DIST";

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

        // Get z value
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double z = args[0].NumericValue;

        // Get cumulative flag
        bool cumulative;
        if (args[1].Type == CellValueType.Boolean)
        {
            cumulative = args[1].BoolValue;
        }
        else if (args[1].Type == CellValueType.Number)
        {
            cumulative = args[1].NumericValue != 0;
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        double result = cumulative ? StatisticalHelper.NormSDist(z) : StatisticalHelper.NormSPdf(z);
        return CellValue.FromNumber(result);
    }
}
