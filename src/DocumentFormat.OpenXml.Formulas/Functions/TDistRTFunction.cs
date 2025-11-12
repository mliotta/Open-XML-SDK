// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the T.DIST.RT function.
/// T.DIST.RT(x, deg_freedom) - returns the right-tailed Student's t-distribution.
/// </summary>
public sealed class TDistRTFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TDistRTFunction Instance = new();

    private TDistRTFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "T.DIST.RT";

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

        // Get x value
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double x = args[0].NumericValue;

        // Get degrees of freedom
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double df = args[1].NumericValue;

        if (df < 1)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            // Right-tailed: P(T > x) = 1 - CDF(x)
            double result = 1.0 - StatisticalHelper.TDistCDF(x, df);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
