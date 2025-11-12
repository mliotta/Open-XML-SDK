// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the F.INV function.
/// F.INV(probability, deg_freedom1, deg_freedom2) - returns the inverse of the F probability distribution.
/// </summary>
public sealed class FInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FInvFunction Instance = new();

    private FInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "F.INV";

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

        if (probability < 0 || probability > 1)
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

        try
        {
            double result = StatisticalHelper.FDistInv(probability, df1, df2);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
