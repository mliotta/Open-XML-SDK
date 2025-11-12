// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FISHERINV function.
/// FISHERINV(y) - returns the inverse of the Fisher transformation.
/// </summary>
public sealed class FisherinvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FisherinvFunction Instance = new();

    private FisherinvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FISHERINV";

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

        double y = args[0].NumericValue;

        // Inverse Fisher transformation: (e^(2y) - 1) / (e^(2y) + 1)
        double e2y = System.Math.Exp(2 * y);
        double result = (e2y - 1) / (e2y + 1);

        return CellValue.FromNumber(result);
    }
}
