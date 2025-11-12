// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the NORM.S.INV function.
/// NORM.S.INV(probability) - returns the inverse of the standard normal cumulative distribution.
/// </summary>
public sealed class NormSInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly NormSInvFunction Instance = new();

    private NormSInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "NORM.S.INV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
        {
            return CellValue.Error("#VALUE!");
        }

        var arg = args[0];

        if (arg.IsError)
        {
            return arg;
        }

        // Get probability
        if (arg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double probability = arg.NumericValue;

        if (probability <= 0 || probability >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        try
        {
            double result = StatisticalHelper.NormSInv(probability);
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
