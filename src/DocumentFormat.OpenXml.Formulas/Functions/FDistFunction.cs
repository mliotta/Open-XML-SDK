// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the F.DIST function.
/// F.DIST(x, deg_freedom1, deg_freedom2, cumulative) - returns the F probability distribution.
/// </summary>
public sealed class FDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FDistFunction Instance = new();

    private FDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "F.DIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
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

        if (x < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get degrees of freedom 1
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double df1 = args[1].NumericValue;

        if (df1 < 1 || df1 > 10000000000)
        {
            return CellValue.Error("#NUM!");
        }

        // Get degrees of freedom 2
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double df2 = args[2].NumericValue;

        if (df2 < 1 || df2 > 10000000000)
        {
            return CellValue.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[3].Type == CellValueType.Boolean)
        {
            cumulative = args[3].BoolValue;
        }
        else if (args[3].Type == CellValueType.Number)
        {
            cumulative = args[3].NumericValue != 0;
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        try
        {
            double result;
            if (cumulative)
            {
                result = StatisticalHelper.FDistCDF(x, df1, df2);
            }
            else
            {
                result = StatisticalHelper.FDistPDF(x, df1, df2);
            }

            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
